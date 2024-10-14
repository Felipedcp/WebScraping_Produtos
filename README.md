# Automação Web - Busca de Ofertas
## Descrição
Esse projeto realiza busca de produtos de forma automatizada nos sites **Google Shopping** e **Buscapé**, tendo como base a planilha **buscas.xlsx**. Os resultados dessa busca são filtrados e armazenados em um arquivo Excel, que é constantemente atualizado com novas buscas, formando um histórico de preços. A cada busca, um e-mail também é enviado com as ofertas encontradas.

## Funcionalidades
- **Busca Automatizada:** Realiza buscas de produtos nos sites **Google Shopping** e **Buscapé**.
- **WebScraping:** Coleta e organiza os dados dos produtos encontrados
- **Histórico de Preços:** Atualiza um arquivo excel com as ofertas encontradas, mantendo um histórico de buscas e de preços ao longo do tempo, pois também armazena data e hora da busca.
- **Notificação por E-mail:** Envia as ofertas encontradas para um endereço de e-mail específico, através do Outlook.

## Bibliotecas Utilizadas
- ***Selenium***
- ***Pandas***
- ***Webdriver Manager***
- ***win32com.client***
- ***datetime***
- ***pythoncom***
- ***pytz***
- ***time***

## Funcionamento (Passo a Passo)

**Passo 1:** Devemos cadastrar na planilha<code>buscas.xlsx</code> todos os produtos que queremos buscar. Devemos preencher os campos **Nome**, **Termos banidos**, **Preço mínimo** e **Preço máximo**. Essa planilha precisa estar na mesma pasta do arquivo <code>main.py</code>. Além disso, um endereço de e-mail deverá ser cadastrado no arquivo <code>main.py</code>

Os campos **Nome** e **Termos banidos** deverão ser preenchidos com as palavras separadas por espaço. O campo **Termos banidos** deverá ser preenchido com palavras que você não gostaria que aparecesse nos produtos encontrados, mas caso você não queira banir nenhuma palavra, é só deixar este campo vazio.

Estrutura da planilha <code>buscas.xlsx</code>
<table>
 <tr><td><b>Nome</b></td><td><b>Termos banidos</b></td><td><b>Preço mínimo</b></td><td><b>Preço máximo</b></td></tr>
 <tr><td>iphone 15 128 gb</td><td>mini watch</td><td>4000</td><td>5000</td></tr>
 <tr><td>rtx 4060</td><td>zotac galax</td><td>2000</td><td>3100</td></tr>
</table>
<br>

**Passo 2:** Utilizando o **Pandas**, o sistema lerá a planilha <code>buscas.xlsx</code> com todos os produtos que deverão ser pesquisados.

**Passo 3:** Com o **Selenium**, o sistema realizará as buscas no **Google Shopping** e **Buscapé** pelos produtos.

**Passo 4:** Para cada produto encontrado é feita uma verificação se ele possui algum termo banido e se ele tem todos os termos da busca no nome. Se ele não possuir nenhum termo banido e possuir todos os termos da busca no nome, e se ele estiver dentro da faixa de preço cadastrada, então ele será adicionado em um **dataframe** contendo informações de **Produto**, **Preço**, **Tipo**, **Data**, **Hora** e **Link**.

**Passo 5:** Os produtos adicionados no **dataframe** serão ordenados em ordem crescente de preço. Se já existir o arquivo <code>tabela_ofertas.xlsx</code> na mesma pasta do arquivo <code>main.py</code>, os resultados dessa busca serão adicionados nessa planilha junto aos resultados anteriores que já se encontram nessa planilha. Caso esse arquivo ainda não exista, ele será criado com os resultados dessa busca com o nome <code>tabela_ofertas.xlsx</code>.

**Passo 6:** Um mensagem contendo uma tabela no formato HTML com todas as ofertas encontradas durante a busca será enviada para o e-mail cadastrado no arquivo <code>main.py</code>. Esse e-mail será enviado através da integração com o Outlook.

Estrutura da planilha <code>tabela_ofertas.xlsx</code>

<table>
 <tr><td><b>Produto</b></td><td><b>Preço</b></td><td><b>Tipo</b></td><td><b>Data</b></td><td><b>Hora</b></td><td><b>Link</b></td></tr>
 
 <tr><td>apple iphone 15 128gb 6gb ram com tela super retina xdr 6.1" - celulares em ...
 </td><td>4599,90
 </td><td>iphone 15 128 gb
 </td><td>10/10/2024
 </td><td>10:23
 </td><td>[Link]
 </td></tr>

 <tr><td>placa de vídeo rtx 4060 ventus 2x white oc msi nvidia geforce, 8gb, gddr6
 </td><td>2020,00
 </td><td>rtx 4060
 </td><td>10/10/2024
 </td><td>10:24
 </td><td>[Link]
 </td></tr>

</table>
<br>

## Requisitos
- **Python 3.8+**
- **Google Chrome (versão mais recente instalada)**
- **ChromeDriver** que será instalado automaticamente pelo **webdriver manager** de acordo com a versão do **Google Chrome** instalada.

## Instalação e Execução

1. **Clone** o repositório

2. Crie e ative um **ambiente virtual**

3. Instale as dependências:
    <b><pre>pip install -r requirements.txt</pre></b>

4. Execute o arquivo <code>main.py</code>:
    <b><pre>python main.py</pre></b>