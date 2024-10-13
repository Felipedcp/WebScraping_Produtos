# Automação Web - Busca de Ofertas
## Descrição
Esse projeto realiza busca de produtos de forma automatizada nos sites **Google Shopping** e **Buscapé**, tendo como base a planilha **buscas.xlsx**. Os resultados dessa busca são filtrados e armazenados em um arquivo Excel, que é constantemente atualizado com novas buscas, formando um histórico de preços. A cada busca, um e-mail também é enviado com as ofertas encontradas.

Estrutura da planilha <code>buscas.xlsx</code>
<table>
 <tr><td><b>Nome</b></td><td><b>Termos banidos</b></td><td><b>Preço mínimo</b></td><td><b>Preço máximo</b></td></tr>
 <tr><td>iphone 15 128 gb</td><td>mini watch</td><td>4000</td><td>5000</td></tr>
 <tr><td>rtx 4060</td><td>zotac galax</td><td>2000</td><td>3100</td></tr>
</table>

## Funcionalidades
- **Busca Automatizada:** Realiza buscas de produtos nos sites **Google Shopping** e **Buscapé**.
- **WebScraping:** Coleta e organiza os dados dos produtos encontrados
- **Histórico de Preços:** Atualiza um arquivo excel com as ofertas encontradas, mantendo um histórico de busca e de preços ao longo do tempo, pois também armazena data e hora da busca.
- **Notificação por E-mail:** Envia as ofertas encontradas para um endereço de e-mail específico, através do Outlook.