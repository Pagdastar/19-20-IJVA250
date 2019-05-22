package com.example.demo.controller;

import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.entity.LigneFacture;
import com.example.demo.repository.ClientRepository;
import com.example.demo.repository.FactureRepository;
import com.example.demo.service.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.PrintWriter;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * Controller pour réaliser les exports.
 */
@Controller
@RequestMapping("/")
public class ExportController {

    @Autowired
    private ClientService clientService;

    @Autowired
    private  FactureService factureService;

    @GetMapping("/clients/csv")
    public void clientsCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("text/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.csv\"");
        PrintWriter writer = response.getWriter();
        List<Client> allClients = clientService.findAllClients();
        LocalDate now = LocalDate.now();
        writer.println("Id Client;Nom;Prénom;Date de Naissance;Âge");

        for (Client client : allClients) {

            writer.println(client.getId() + ";"
                    + client.getNom() + ";"
                    + client.getPrenom() + ";"
                    + client.getDateNaissance().format(DateTimeFormatter.ofPattern("dd/MM/YYYY")) + ";"
                    + (now.getYear() - client.getDateNaissance().getYear()) + " ans");
        }
    }

    @GetMapping("/clients/xlsx")
    public void clientsXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename=\"clients.xlsx\"");
        LocalDate now = LocalDate.now();

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Clients");
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("ID Client");
        headerRow.createCell(1).setCellValue("Prénom");
        headerRow.createCell(2).setCellValue("Nom");
        headerRow.createCell(3).setCellValue("Date de naissance");
        headerRow.createCell(4).setCellValue("Âge");

        List<Client> allClients = clientService.findAllClients();

        for (Client client: allClients
        ) {
            Row row = sheet.createRow(sheet.getLastRowNum()+1);
            row.createCell(0).setCellValue(client.getId());
            row.createCell(2).setCellValue(client.getNom());
            row.createCell(1).setCellValue(client.getPrenom());
            row.createCell(3).setCellValue(client.getDateNaissance().toString());
            row.createCell(4).setCellValue(now.getYear() - client.getDateNaissance().getYear()  + " ans");


        }

        workbook.write(response.getOutputStream());
        workbook.close();
    }

    @GetMapping("/clients/{id}/factures/xlsx")
    public void facturesClientsXLSX(@PathVariable("id") Long clientId, HttpServletRequest request, HttpServletResponse response) throws IOException{
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename=\"factures-client-" + clientId + ".xlsx\"");
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Factures");
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("ID Facture");
        headerRow.createCell(1).setCellValue("Total");

        List<Facture> allFactures = factureService.findByClientId(clientId);

        generateFactureWorkbook(response, workbook, sheet, allFactures);
    }

    private void generateFactureWorkbook(HttpServletResponse response, Workbook workbook, Sheet sheet, List<Facture> allFactures) throws IOException {
        for (Facture facture: allFactures) {
            Row row = sheet.createRow(sheet.getLastRowNum()+1);
            row.createCell(0).setCellValue(facture.getId());
            row.createCell(1).setCellValue(facture.getTotal());
        }

        workbook.write(response.getOutputStream());
        workbook.close();
    }

    @GetMapping("/factures/xlsx")
    public void facturesXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException{
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename=\"factures.xlsx\"");

        List<Client> allClients = clientService.findAllClients();

        // Création du fichier
        Workbook workbook = new XSSFWorkbook();

        for (Client client: allClients) {
            // Création onglet client
            Sheet clientSheet = workbook.createSheet(client.getNom());

            Row clientRow = clientSheet.createRow(0);

            Cell cellPrenom = clientRow.createCell(0);
            cellPrenom.setCellValue("Prénom");
            Cell cellNom = clientRow.createCell(1);
            cellNom.setCellValue("Nom");

            Row rowClient = clientSheet.createRow(1);

            Cell prenom = rowClient.createCell(0);
            prenom.setCellValue(client.getPrenom());
            Cell nom = rowClient.createCell(1);
            nom.setCellValue(client.getNom());

            // Création onglet facture par client

            List<Facture> facturesClient = factureService.findByClientId(client.getId());

            for (Facture facture: facturesClient) {

            // Création feuille facture
                Sheet factureSheet = workbook.createSheet("Facture n°" + facture.getId());

                Set<LigneFacture> ligneFactures = facture.getLigneFactures();

                Row headerFacture = factureSheet.createRow(0);

                Cell cellLibelle = headerFacture.createCell(0);
                cellLibelle.setCellValue("Nom de l'article");

                Cell cellQuantite = headerFacture.createCell(1);
                cellQuantite.setCellValue("Quantité");

                Cell cellPrix = headerFacture.createCell(2);
                cellPrix.setCellValue("Prix unitaire");

                Cell cellTotalLigne = headerFacture.createCell(3);
                cellTotalLigne.setCellValue("Prix total de la ligne");

                // Ajout des donnees de LigneFacture dans la feuille Facture

                int iRow = 1;
                for(LigneFacture ligneFacture : facture.getLigneFactures()){
                    Row headerArticle = factureSheet.createRow(iRow);

                    Cell cellNomArticle = headerArticle.createCell(0);
                    cellNomArticle.setCellValue(ligneFacture.getArticle().getLibelle());

                    Cell cellQuantiteArticle = headerArticle.createCell(1);
                    cellQuantiteArticle.setCellValue(ligneFacture.getQuantite());

                    Cell cellPrixArticle = headerArticle.createCell(2);
                    cellPrixArticle.setCellValue(ligneFacture.getArticle().getPrix());

                    Cell cellTotalLigneArticle = headerArticle.createCell(3);
                    cellTotalLigneArticle.setCellValue(ligneFacture.getSousTotal());

                    iRow +=1;
                }

                // Ajustement auto des colonnes
                factureSheet.autoSizeColumn(0);
                factureSheet.autoSizeColumn(1);
                factureSheet.autoSizeColumn(2);
                factureSheet.autoSizeColumn(3);
                factureSheet.autoSizeColumn(0);
                factureSheet.autoSizeColumn(1);

                // Ajout de TOTAL dans la facture

                // Style de la cellule TOTAL
                CellStyle styleLibelle = workbook.createCellStyle();
                styleLibelle.setBorderBottom(BorderStyle.MEDIUM);
                styleLibelle.setBorderTop(BorderStyle.MEDIUM);
                styleLibelle.setBorderLeft(BorderStyle.MEDIUM);
                styleLibelle.setBottomBorderColor(IndexedColors.RED.getIndex());
                styleLibelle.setTopBorderColor(IndexedColors.RED.getIndex());
                styleLibelle.setLeftBorderColor(IndexedColors.RED.getIndex());

                Font font = workbook.createFont();
                font.setBold(true);
                font.setColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
                styleLibelle.setFont(font);

                // Style de la cellule TOTAL GÉNÉRAL
                CellStyle styleValeur = workbook.createCellStyle();
                styleValeur.setBorderBottom(BorderStyle.MEDIUM);
                styleValeur.setBorderTop(BorderStyle.MEDIUM);
                styleValeur.setBorderRight(BorderStyle.MEDIUM);
                styleValeur.setBottomBorderColor(IndexedColors.RED.getIndex());
                styleValeur.setTopBorderColor(IndexedColors.RED.getIndex());
                styleValeur.setRightBorderColor(IndexedColors.RED.getIndex());

                Font total = workbook.createFont();
                total.setBold(true);
                total.setColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
                styleValeur.setFont(total);


                int iTotal = iRow+1;
                Row headerTotal = factureSheet.createRow(iTotal);
                Cell cellTotalLibelle = headerTotal.createCell(0);
                Cell cellTotalLibelle1 = headerTotal.createCell(1);
                Cell cellTotalLibelle2 = headerTotal.createCell(2);
                cellTotalLibelle.setCellValue("TOTAL GÉNÉRAL");

                cellTotalLibelle.setCellStyle(styleLibelle);
                cellTotalLibelle1.setCellStyle(styleLibelle);
                cellTotalLibelle2.setCellStyle(styleLibelle);

                // Fusion de cellules pour le libellé Total
                factureSheet.addMergedRegion(new CellRangeAddress((iTotal),(iTotal),0,2));

                Cell cellTotal = headerTotal.createCell(3);
                cellTotal.setCellValue(facture.getTotal());
                cellTotal.setCellStyle(styleValeur);
            }
        }
        workbook.write(response.getOutputStream());
        workbook.close();
    }



}