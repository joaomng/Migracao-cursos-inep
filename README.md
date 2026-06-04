# Migracao-cursos-inep
Repositório para a automatização da migração dos cursos para o censo da educação superior

## Dependências:

- pandas <br/>

   pip install pandas 

## Como utilizar:

<b> 1. Instalar python: </b> <br/>
    No windows, pesquisar "python" na barra de pesquisa e instalar conforme a página da microsoft store que será aberta. <br/>

<b> 2. Instalar dependências </b> <br/>
    No windows, pesquisar "cmd" na barra de pesquisa. Apertando Enter será aberto o prompt de comando. <br/>
    No prompt, digitar: 
      
      pip install pandas

<b> 3. Executar o programa </b> <br/>
  3.1. Fazer o download deste repositório; <br/>
  3.2. Extrair o repositório; <br/>
  3.3. Abrir o cmd e ir para a pasta do repositório. O comando "cd" abre uma pasta, por exemplo, o comando: <br/>
  
       cd Documentos 
  
  abre a pasta Documentos. O comando <br/> 
  
       cd Documentos\Pasta_1 
       
  abre a pasta Pasta_1, que está dentro da pasta documentos <br/>
  
  <b> 3.4. Digitar o seguinte comando </b> <br/>
  
        python migracao_cursos.py 

  O programa iniciará a execução

  <b> 3.5. Siga as instruções na tela </b> <br/>
  Será solicitado o nome das planilhas, é preciso digitá-los corretamente, incluindo maiúsculas e minúsculas. <br/>
  O mesmo vale para as abas das planilhas.
        

## Obs:
Esse repositório também tem um notebook .ipynb para uma execução interativa dos scripts de migração, possibilitando mudar os parâmetros das funções utilizadas.


## Importante:
Só funciona com planilhas que estejam formatadas conforme as planilhas de exemplo (no caso da planilha de cursos, por exemplo, <br/> com o cabeçalho na SEGUNDA linha e as colunas na mesma ordem até a coluna L (as demais não são usadas diretamente pelo script)), e só para cursos 100% presenciais e que TIVERAM ALUNOS INSCRITOS NO ANO DE REFERÊNCIA
