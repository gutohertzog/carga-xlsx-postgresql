# carga-xlsx-postgresql

Esse positório sugiu surgiu a partir de uma demanda do trabalho.

Necessidade : realizar a atualização dos dados de uma tabela no banco de dados PostgreSQL a paritr de um arquivo no formato xlsx.

Numa primeira versão, o script Python apenas irá buscar os dados no arquivo xlsx e então gerar as queries,
que serão executas no pgAdmin.

Para realizar isso, é necessário usar algumas bibliotecas externas do Python.

Bibliotecas :
- openpyxl : usado para abrir e carregar os dados do arquivo xlsx;

PS.: por motivos de segurança, nomes de tabelas e campos serão generalizados.

