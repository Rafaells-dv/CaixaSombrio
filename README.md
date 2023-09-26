Sistema de Registro de Vendas

Este é um sistema de controle de vendas simples desenvolvido em Python, projetado para registrar vendas em um arquivo Excel para facilitar a análise de dados.
Desenvolvido para um motoclube local.

Funcionalidades:

Registro de Vendas: O sistema permite registrar vendas de produtos ou serviços, produtos vendidos e valor total da venda. As vendas são automaticamente registradas em um banco de dados.

Geração de Arquivo Excel: Cada venda é adicionada a um banco de dados que pode ser exportado para um Excel servindo como um registro detalhado de todas as transações. Isso facilita a análise de dados e o acompanhamento das vendas ao longo do tempo.

Requisitos:

- SQLite ODBC Driver for Win64
- Arquivo Excel com as colunas [CODIGOPRODUTO, VALOR, PRODUTO](Utilizar codigoprodutos.xlsx presente no repositório)