# Autodoc - Gerador de Documentação de Banco de Dados

O **Autodoc** é uma ferramenta automatizada desenvolvida para gerar documentação completa de bancos de dados PostgreSQL. Ele cria um relatório consolidado em PDF contendo um diagrama Entidade-Relacionamento (ER) e um Dicionário de Dados detalhado, além de gerar os arquivos fonte (Excel e Word) separadamente.

## Requisitos do Sistema

* **Sistema Operacional**: Linux (testado em Ubuntu/Debian).
* **Python**: 3.8 ou superior.
* **Dependências Externas**:
  * `LibreOffice`: Para conversão de documentos `.docx` e `.xlsx` para PDF.
  * `Poppler-utils` (pdfunite): Para mesclagem dos arquivos PDF.
  * `Graphviz`: Para renderização dos diagramas ER.

## Instalação

1. **Clone o repositório** (se aplicável) ou navegue até o diretório do projeto.

2. **Instale as dependências do Sistema Operacional**:
    Execute o script auxiliar para instalar LibreOffice, Poppler e Graphviz:

    ```bash
    ./install_dependencies.sh
    ```

3. **Configure o Ambiente Python**:
    Recomenda-se o uso de um ambiente virtual (Conda ou Venv).

    ```bash
    # Exemplo com Conda
    conda create -n autodoc python=3.13
    conda activate autodoc
    
    # Instale os pacotes Python
    pip install -r requirements.txt
    ```

## Configuração

Antes de executar, você pode personalizar a geração:

* **`details.txt`**: Define o Nome do Projeto, Título e Descrição do relatório.

    ```text
    Project: NOME DO PROJETO
    Title: Título do Relatório
    Description: Descrição breve do escopo do banco de dados.
    ```

* **`excluded_tables.txt`**: Lista de tabelas que devem ser ignoradas no diagrama e no dicionário (separadas por vírgula).

    ```text
    django_migrations, auth_user, ...
    ```

* **`model.docx`**: Um documento Word que serve como template para a capa/introdução.

## Uso

Para gerar a documentação, execute o script principal:

```bash
python generate_er_doc.py
```

O script solicitará as credenciais de conexão com o banco de dados PostgreSQL:

* Host (padrão: localhost)
* Porta (padrão: 5432)
* Usuário
* Senha
* Nome do Banco de Dados

### Saídas Geradas

Ao final da execução, os seguintes arquivos serão criados no diretório do projeto (com timestamp no nome):

1. **`Autodoc_<NOME_DB>_<TIMESTAMP>.pdf`**: O relatório final completo e unificado.
2. **`Autodoc_DataDictionary_<TIMESTAMP>.xlsx`**: O Dicionário de Dados em formato Excel editável.
3. Arquivos temporários (`temp_*.pdf`, `temp_*.docx`) são gerados durante o processo e removidos automaticamente (exceto em caso de erro).

## Solução de Problemas

* **Erro "soffice not found"**: Certifique-se de que o LibreOffice está instalado e acessível no PATH.
* **Erro de Conexão com Banco**: Verifique se o PostgreSQL está rodando e se as credenciais (usuário/senha/banco) estão corretas.
* **Erro ao gerar PDF**: Verifique se nenhum arquivo com o mesmo nome está aberto em outro programa.
