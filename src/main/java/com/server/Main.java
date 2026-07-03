package com.server;

import java.net.InetSocketAddress;
import java.util.concurrent.Executors;
import java.util.logging.Logger;
import com.sun.net.httpserver.HttpServer;

public class Main {

    private static final Logger LOG = Logger.getLogger(Main.class.getName());

    public static void main(String[] args) throws Exception {
        int port = 8080;
        if (args.length > 0) {
            port = Integer.parseInt(args[0]);
        }

        HttpServer server = HttpServer.create(new InetSocketAddress("127.0.0.1", port), 0);
        server.createContext("/api/consolidaconsuntivi", new ConsolidaHandler());
        server.createContext("/api/test-download", new TestDownloadHandler());
        server.createContext("/api/health", exchange -> {
            String body = "{\"status\":\"ok\",\"version\":\"1.1.0\"}";
            exchange.getResponseHeaders().add("Content-Type", "application/json");
            exchange.sendResponseHeaders(200, body.length());
            exchange.getResponseBody().write(body.getBytes());
            exchange.close();
        });
        server.setExecutor(Executors.newFixedThreadPool(4));
        server.start();

        LOG.info("Server avviato su http://127.0.0.1:" + port + "/api/consolidaconsuntivi");
    }
}
