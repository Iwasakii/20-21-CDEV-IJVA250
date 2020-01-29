package com.example.demo.controller;

import com.example.demo.entity.Article;
import com.example.demo.entity.Client;
import com.example.demo.entity.Facture;
import com.example.demo.service.impl.ClientServiceImpl;
import com.example.demo.service.ArticleService;
import com.example.demo.service.FactureService;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.servlet.ModelAndView;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.List;


/**
 * Controller principale pour affichage des clients / factures sur la page d'acceuil.
 */
@Controller
public class HomeController {

    private ArticleService articleService;
    private ClientServiceImpl clientServiceImpl;
    private FactureService factureService;

    public HomeController(ArticleService articleService, ClientServiceImpl clientService, FactureService factureService) {
        this.articleService = articleService;
        this.clientServiceImpl = clientService;
        this.factureService = factureService;
    }

    @GetMapping("/")
    public ModelAndView home() {
        ModelAndView modelAndView = new ModelAndView("home");

        List<Article> articles = articleService.findAll();
        modelAndView.addObject("articles", articles);

        List<Client> toto = clientServiceImpl.findAllClients();
        modelAndView.addObject("clients", toto);

        List<Facture> factures = factureService.findAllFactures();
        modelAndView.addObject("factures", factures);

        return modelAndView;
    }

    @GetMapping("articles/csv")
    public void articleCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("test/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"export-articles.csv\"");
        PrintWriter printWriter = response.getWriter();
        List<Article> articles = articleService.findAll();
        for (Article article : articles) {
            printWriter.print(article.getLibelle() + ";");
            printWriter.println(article.getPrix());
        }
    }

    @GetMapping("clients/csv")
    public void clientCSV(HttpServletRequest request, HttpServletResponse response) throws IOException {
        response.setContentType("test/csv");
        response.setHeader("Content-Disposition", "attachment; filename=\"export-clients.csv\"");
        PrintWriter printWriter = response.getWriter();
        List<Client> clients = clientServiceImpl.findAllClients();
        for (Client client : clients) {
            printWriter.print(client.getNom() + ";");
            printWriter.println(client.getPrenom());
        }
    }

    @GetMapping("articles/xlsx")
    public void articleXLSX(HttpServletRequest request, HttpServletResponse response) throws IOException {
        ServletOutputStream os = response.getOutputStream();
        response.setContentType("test/xlsx");
        response.setHeader("Content-Disposition", "attachment; filename=\"export-article.xlsx\"");
        List<Article> articles = articleService.findAll();
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Clients");
        Row headerRow = sheet.createRow(0);
        Cell cellPrenom = headerRow.createCell(0);
        cellPrenom.setCellValue("Prénom");
        for(int i = 0; i < articles.size(); i++) {
            Row row = sheet.createRow(i);
            Cell cellP = row.createCell(i);
            Cell cellN = row.createCell(i);
            cellP.setCellValue(articles.get(i).getLibelle());
            cellN.setCellValue(articles.get(i).getPrix());
        }
        workbook.write(os);
        workbook.close();
    }

}

