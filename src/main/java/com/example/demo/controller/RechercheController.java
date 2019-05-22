package com.example.demo.controller;

import com.example.demo.entity.Article;
import com.example.demo.service.ArticleService;
import com.example.demo.service.FactureService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.List;

/**
 * Contrôleur pour réaliser les exports.
 */
@Controller
@RequestMapping("/")
public class RechercheController {

    @Autowired
    private ArticleService articleService;

    @GetMapping("/recherche")
    public void rechercher(
            @RequestParam String query,
            HttpServletResponse response
    ) throws IOException {
        List<Article> articles = articleService.find(query);
        for (Article article : articles) {
            response.getWriter().println(article.getLibelle());
            response.getWriter().println("<BR/>");
        }
    }

}