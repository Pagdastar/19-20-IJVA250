package com.exemple.demo.entity;

import com.example.demo.entity.Article;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import org.assertj.core.api.Assertions;
import org.junit.Assert;
import org.junit.Test;

import java.util.HashSet;

import static org.junit.Assert.assertEquals;

public class FactureTest {

    @Test
    public void getTotal_factureVideSansArticleRetourneZero() {
        // GIVEN
        Facture facture = new Facture();
        facture.setLigneFactures(new HashSet<>());
        // WHEN
        Double total = facture.getTotal();
        // THEN
        Assertions.assertThat(total).isEqualTo(0.0);
    }

    @Test
    public void getTotal_factureAvecUnArticleRetournePrixUnitaireDeLArticle() {
        // GIVEN
        Article article = new Article();
        article.setPrix(3.99);

        LigneFacture ligneFacture = new LigneFacture();
        ligneFacture.setArticle(article);
        ligneFacture.setQuantite(1);

        Facture facture = new Facture();
        HashSet<LigneFacture> ligneFactures = new HashSet<>();
        ligneFactures.add(ligneFacture);
        facture.setLigneFactures(ligneFactures);

        // WHEN
        Double total = facture.getTotal();
        // THEN
        Assertions.assertThat(total).isEqualTo(3.99);
    }

    @Test
    public void getTotal_factureAvecDeuxFoisLeMemeArticleRetournePrixUnitaireDeLArticleFoisDeux() {
        // GIVEN
        Article article = new Article();
        article.setPrix(3.99);

        LigneFacture ligneFacture = new LigneFacture();
        ligneFacture.setArticle(article);
        ligneFacture.setQuantite(2);

        Facture facture = new Facture();
        HashSet<LigneFacture> ligneFactures = new HashSet<>();
        ligneFactures.add(ligneFacture);
        facture.setLigneFactures(ligneFactures);

        // WHEN
        Double total = facture.getTotal();
        // THEN
        Assertions.assertThat(total).isEqualTo(7.98);
    }
}
