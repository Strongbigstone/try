package org.example;

import cn.hutool.core.io.FileUtil;
import cn.hutool.core.util.StrUtil;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.IOException;
import java.net.URL;
import java.util.HashSet;
import java.util.Set;
import java.util.regex.Pattern;
import java.util.stream.IntStream;

public class Crawler {
    private final Pattern urlPattern = Pattern.compile("^https?://.+$");
    private Set<String> visited = new HashSet<>();
    private String filePath;

    public Crawler(String url, String filePath) {
        if (FileUtil.isDirectory(filePath)) FileUtil.mkdir(filePath);
        this.filePath = filePath;
        dfs(url);
        System.out.println("--------------> end");
    }

    public void dfs(String url) {
        if (!set(url)) return;
        System.out.println("visit to " + url);
        try {
            Document doc = Jsoup.connect(url).get();
            Elements img = doc.getElementsByTag("img");
            downloadImg(img);
            for (Element link : doc.select("a[href]")) {
                String nextUrl = link.attr("href");
                dfs(nextUrl);
            }
        } catch (Exception ignored) {
            System.out.println("--------------> is failed");
        }
    }

    public boolean set(String url) {
        if (urlPattern.matcher(url).find()) return visited.add(url);
        else return false;
    }

    public void downloadImg(Elements img) {
        IntStream.range(0, img.size())
                .forEach(i -> {
                    String src = img.get(i).attr("src");
                    if (StrUtil.startWith(src, "//")) src = String.format("http:%s", src);
                        try {
                            URL url = new URL(src);
                            FileUtil.writeFromStream(url.openStream(), appendPath(filePath));
                        } catch (IOException ignored) {
                        }
                });
    }

    private static String appendPath(String filePath) {
        return filePath + "//img_" + System.currentTimeMillis() + ".jpg";
    }

    public static void main(String[] args) {
        new Crawler("http://www.baidu.com", "D://testDownload");
    }
}
