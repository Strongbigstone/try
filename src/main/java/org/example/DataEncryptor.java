import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.core.type.TypeReference;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Component;
import org.springframework.util.FileSystemUtils;

import javax.annotation.PostConstruct;
import javax.crypto.Cipher;
import javax.crypto.spec.IvParameterSpec;
import javax.crypto.spec.SecretKeySpec;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.locks.ReentrantLock;
import org.apache.commons.codec.binary.Base64;

@
public class DataEncryptor {

    private static final Logger logger = LoggerFactory.getLogger(DataEncryptor.class);

    // 密钥管理 - 实际生产应使用KMS
    @Value("${encryption.secret.key}")
    private String secretKeyStr;

    @Value("${encryption.iv}")
    private String ivStr;

    @Value("${encryption.config.file}")
    private String configFilePath;

    @Value("${encryption.checkpoint.file}")
    private String checkpointFilePath;

    @Value("${encryption.batch.size:1000}")
    private int batchSize;

    // 同步锁用于JSON文件写入
    private final ReentrantLock checkpointLock = new ReentrantLock();

    @Autowired
    private JdbcTemplate jdbcTemplate;

    // 表加密配置 (tableName -> TableEncryptionConfig)
    private Map<String, TableEncryptionConfig> encryptionConfigs = new HashMap<>();

    // 断点记录 (tableName -> lastProcessedId)
    private Map<String, Object> encryptionCheckpoints = new ConcurrentHashMap<>();

    // 加密组件
    private SecretKeySpec secretKey;
    private IvParameterSpec iv;

    // JSON处理器
    private ObjectMapper objectMapper = new ObjectMapper();

    @PostConstruct
    public void init() throws Exception {
        // 加载加密配置
        loadEncryptionConfig();

        // 初始化加密密钥
        initEncryptionKey();

        // 加载断点记录
        loadEncryptionCheckpoints();
    }

    private void loadEncryptionConfig() {
        try {
            Path path = Paths.get(configFilePath);
            if (Files.exists(path)) {
                encryptionConfigs = objectMapper.readValue(
                        Files.readAllBytes(path),
                        new TypeReference<Map<String, TableEncryptionConfig>>(){}
                );
                logger.info("加密配置加载成功: {} 张表", encryptionConfigs.size());
            } else {
                logger.error("加密配置文件不存在: {}", configFilePath);
                // 创建默认配置示例
                createDefaultConfig();
            }
        } catch (Exception e) {
            logger.error("加载加密配置失败", e);
            throw new RuntimeException("加载加密配置失败", e);
        }
    }

    private void createDefaultConfig() {
        // 默认配置示例
        encryptionConfigs = new HashMap<>();
        encryptionConfigs.put("users", new TableEncryptionConfig(
                "users", "user_id", Arrays.asList("email", "phone", "ssn")
        ));
        encryptionConfigs.put("orders", new TableEncryptionConfig(
                "orders", "order_id", Arrays.asList("credit_card")
        ));

        // 保存默认配置
        saveEncryptionConfig();
    }

    private void saveEncryptionConfig() {
        try {
            Path path = Paths.get(configFilePath);
            Files.createDirectories(path.getParent());
            Files.write(path, objectMapper.writeValueAsBytes(encryptionConfigs));
            logger.info("加密配置已保存到: {}", configFilePath);
        } catch (Exception e) {
            logger.error("保存加密配置失败", e);
        }
    }

    private void initEncryptionKey() {
        // 使用AES-256-CBC模式，带初始化向量
        secretKey = new SecretKeySpec(
                secretKeyStr.getBytes(StandardCharsets.UTF_8), "AES");
        iv = new IvParameterSpec(ivStr.getBytes(StandardCharsets.UTF_8));
    }

    private void loadEncryptionCheckpoints() {
        try {
            Path path = Paths.get(checkpointFilePath);
            if (Files.exists(path)) {
                encryptionCheckpoints = objectMapper.readValue(
                        Files.readAllBytes(path),
                        new TypeReference<Map<String, Object>>(){}
                );
                logger.info("断点信息加载成功: {} 个断点", encryptionCheckpoints.size());
            } else {
                logger.warn("断点文件不存在，创建新文件: {}", checkpointFilePath);
                Files.createDirectories(path.getParent());
                Files.createFile(path);
                encryptionCheckpoints = new HashMap<>();
                saveEncryptionCheckpoints();
            }
        } catch (Exception e) {
            logger.error("加载断点信息失败", e);
            throw new RuntimeException("加载断点信息失败", e);
        }
    }

    private void saveEncryptionCheckpoints() {
        checkpointLock.lock();
        try {
            Path path = Paths.get(checkpointFilePath);
            Files.write(path, objectMapper.writeValueAsBytes(encryptionCheckpoints));
            logger.debug("断点信息已保存");
        } catch (Exception e) {
            logger.error("保存断点信息失败", e);
        } finally {
            checkpointLock.unlock();
        }
    }

    @Scheduled(cron = "${encryption.schedule.cron:0 0 2 * * ?}")
    public void encryptSensitiveData() {
        logger.info("=== 开始数据加密任务 ===");

        // 更新配置确保最新
        loadEncryptionConfig();

        encryptionConfigs.forEach((configKey, config) -> {
            // 处理每个表的加密
            encryptTableData(config);
        });

        logger.info("=== 数据加密任务完成 ===");
    }

    private void encryptTableData(TableEncryptionConfig config) {
        logger.info("开始处理表: {}", config.getTableName());

        String tableName = config.getTableName();
        String primaryKey = config.getPrimaryKey();

        // 获取当前断点（默认为0或NULL）
        Object startPoint = encryptionCheckpoints.getOrDefault(tableName, "0");
        if ("NULL".equals(startPoint)) startPoint = null;

        int processed = 0;
        boolean completed = false;

        while (!completed) {
            try {
                // 查询批处理数据
                List<Map<String, Object>> rows = fetchDataBatch(config, startPoint);
                if (rows.isEmpty()) {
                    completed = true;
                    break;
                }

                // 处理数据加密
                boolean hasUpdate = processBatchRows(config, rows);
                if (hasUpdate) {
                    processed += rows.size();
                    logger.info("表 {} 已处理 {}/{} 条记录",
                            tableName, processed, "-");

                    // 更新断点并保存
                    Object lastId = rows.get(rows.size() - 1).get(primaryKey);
                    encryptionCheckpoints.put(tableName, lastId != null ? lastId : "NULL");
                    saveEncryptionCheckpoints();
                    startPoint = lastId;
                } else {
                    completed = true;
                }
            } catch (Exception e) {
                logger.error("处理表 {} 时发生错误", tableName, e);
                break;
            }
        }

        // 处理完成后清除断点
        if (completed) {
            encryptionCheckpoints.remove(tableName);
            saveEncryptionCheckpoints();
            logger.info("表 {} 已完成加密处理，共处理 {} 条记录", tableName, processed);
        }
    }

    private List<Map<String, Object>> fetchDataBatch(TableEncryptionConfig config, Object startPoint) {
        String tableName = config.getTableName();
        String primaryKey = config.getPrimaryKey();

        // 构建查询SQL
        String query = buildBatchQuery(config, startPoint);

        // 执行查询
        return jdbcTemplate.queryForList(query, startPoint, batchSize);
    }

    private String buildBatchQuery(TableEncryptionConfig config, Object startPoint) {
        String tableName = config.getTableName();
        String primaryKey = config.getPrimaryKey();
        List<String> columns = config.getColumnsToEncrypt();

        // 基本查询
        StringBuilder sql = new StringBuilder("SELECT ")
                .append(primaryKey).append(", ")
                .append(String.join(", ", columns))
                .append(" FROM ").append(tableName);

        // 添加断点过滤
        if (startPoint != null && !"NULL".equals(startPoint)) {
            sql.append(" WHERE ").append(primaryKey).append(" > ?");
        } else if ("NULL".equals(startPoint)) {
            // NULL值处理
            sql.append(" WHERE ").append(primaryKey).append(" IS NOT NULL");
        }

        // 添加排序
        sql.append(" ORDER BY ").append(primaryKey);

        // 添加限制
        sql.append(" LIMIT ?");

        return sql.toString();
    }

    private boolean processBatchRows(TableEncryptionConfig config, List<Map<String, Object>> rows) {
        if (rows.isEmpty()) return false;

        String primaryKey = config.getPrimaryKey();
        List<String> updateStatements = new ArrayList<>();
        List<Object[]> batchParams = new ArrayList<>();

        // 准备批量更新
        for (Map<String, Object> row : rows) {
            Object idValue = row.get(primaryKey);
            if (idValue == null) continue;

            Map<String, String> encryptedValues = new HashMap<>();

            // 加密需要处理的字段
            for (String column : config.getColumnsToEncrypt()) {
                Object value = row.get(column);
                if (value != null && !value.toString().isEmpty()) {
                    encryptedValues.put(column, encryptData(value.toString()));
                }
            }

            // 创建更新语句
            if (!encryptedValues.isEmpty()) {
                batchParams.add(buildUpdateParams(config, encryptedValues, idValue));
            }
        }

        // 执行批量更新
        if (!batchParams.isEmpty()) {
            executeBatchUpdate(config, batchParams);
            return true;
        }

        return false;
    }

    private Object[] buildUpdateParams(TableEncryptionConfig config, Map<String, String> encryptedValues, Object idValue) {
        // 参数顺序: 加密字段1, 加密字段2, ..., 主键值
        List<Object> params = new ArrayList<>(encryptedValues.values());
        params.add(idValue);
        return params.toArray();
    }

    private void executeBatchUpdate(TableEncryptionConfig config, List<Object[]> batchParams) {
        String tableName = config.getTableName();
        String primaryKey = config.getPrimaryKey();
        List<String> columns = config.getColumnsToEncrypt();

        // 构建UPDATE语句
        StringBuilder sql = new StringBuilder("UPDATE ")
                .append(tableName)
                .append(" SET ");

        // 添加所有需要更新的字段
        for (int i = 0; i < columns.size(); i++) {
            if (i > 0) sql.append(", ");
            sql.append(columns.get(i)).append(" = ?");
        }

        // 添加WHERE条件
        sql.append(" WHERE ").append(primaryKey).append(" = ?");

        // 执行批量更新
        jdbcTemplate.batchUpdate(sql.toString(), batchParams);
    }

    public String encryptData(String plaintext) {
        try {
            Cipher cipher = Cipher.getInstance("AES/CBC/PKCS5Padding");
            cipher.init(Cipher.ENCRYPT_MODE, secretKey, iv);
            byte[] encryptedBytes = cipher.doFinal(plaintext.getBytes(StandardCharsets.UTF_8));
            return Base64.encodeBase64String(encryptedBytes);
        } catch (Exception e) {
            throw new RuntimeException("加密失败: " + e.getMessage(), e);
        }
    }

    public String decryptData(String ciphertext) {
        try {
            Cipher cipher = Cipher.getInstance("AES/CBC/PKCS5Padding");
            cipher.init(Cipher.DECRYPT_MODE, secretKey, iv);
            byte[] decodedBytes = Base64.decodeBase64(ciphertext);
            byte[] decryptedBytes = cipher.doFinal(decodedBytes);
            return new String(decryptedBytes, StandardCharsets.UTF_8);
        } catch (Exception e) {
            throw new RuntimeException("解密失败: " + e.getMessage(), e);
        }
    }

    // 用于JSON序列化的配置类
    public static class TableEncryptionConfig {
        private String tableName;
        private String primaryKey;
        private List<String> columnsToEncrypt;

        // 默认构造方法用于JSON反序列化
        public TableEncryptionConfig() {}

        public TableEncryptionConfig(String tableName, String primaryKey, List<String> columnsToEncrypt) {
            this.tableName = tableName;
            this.primaryKey = primaryKey;
            this.columnsToEncrypt = columnsToEncrypt;
        }

        // Getters and setters
        public String getTableName() { return tableName; }
        public void setTableName(String tableName) { this.tableName = tableName; }

        public String getPrimaryKey() { return primaryKey; }
        public void setPrimaryKey(String primaryKey) { this.primaryKey = primaryKey; }

        public List<String> getColumnsToEncrypt() { return columnsToEncrypt; }
        public void setColumnsToEncrypt(List<String> columnsToEncrypt) {
            this.columnsToEncrypt = columnsToEncrypt;
        }
    }
}