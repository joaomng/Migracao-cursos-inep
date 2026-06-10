

#tem que fazer o pip install pandas

import pandas as pd
import re
#import unicodedata

def acessibilidade_cursos(nome_origem, aba_origem = "Acessibilidade_Cursos"):
  '''A partir da planilha de acessibilidade <nome_origem>, formatada
  como a planilha de 2024, gera um dicionário onde cada chave(string) é o código do curso,
  e cada valor é um outro dicionário: cada chave(string) é cada recurso de acessibilidade,
  conforme será discriminado abaixo, e cada valor(int) é 0 ou 1 se o curso oferece aquele recurso '''



  pcd_por_curso = dict() #a chave é o código do curso (string)
                         #e o valor é um outro dicionario com valores 0 e 1 representando a acessibilidade
                         #com as seguintes chaves:
#mat_braille, info_acessivel, mat_pedagogico_tatil, tradutor_libras, mat_didatico_libras,
#mat_didatico_impresso_acessivel, mat_audio, mat_impresso_ampliado, acessibilidade_comunicacao,
#guia_interprete, libras_ofertada, mat_digital_acessivel

#a planilha de acessibilidade precisa estar na ordem acima!

  xls = pd.ExcelFile(nome_origem)
  df1 = pd.read_excel(xls, aba_origem, header=0)

  n_linhas, n_col = df1.shape


  for i in range(n_linhas):
    linha = df1.iloc[i]

    dict_curso = dict()

    cod_curso = str(linha.iloc[0])

    mat_braille = str(linha.iloc[5])
    dict_curso['mat_braille'] = 0
    if((mat_braille=="X") or (mat_braille == "x")):
      dict_curso['mat_braille'] = 1


    info_acessivel = str(linha.iloc[6])
    dict_curso['info_acessivel'] = 0
    if((info_acessivel=="X") or (info_acessivel == "x")):
      dict_curso['info_acessivel'] = 1



    #mat_pedagogico_tatil
    mat_pedagogico_tatil = str(linha.iloc[7])
    dict_curso['mat_pedagogico_tatil'] = 0
    if((mat_pedagogico_tatil=="X") or (mat_pedagogico_tatil == "x")):
      dict_curso['mat_pedagogico_tatil'] = 1

    #tradutor_libras
    tradutor_libras = str(linha.iloc[8])
    dict_curso['tradutor_libras'] = 0
    if((tradutor_libras=="X") or (tradutor_libras == "x")):
      dict_curso['tradutor_libras'] = 1



    #mat_didatico_libras
    mat_didatico_libras = str(linha.iloc[9])
    dict_curso['mat_didatico_libras'] = 0
    if((mat_didatico_libras=="X") or (mat_didatico_libras == "x")):
      dict_curso['mat_didatico_libras'] = 1


    #mat_didatico_impresso_acessivel
    mat_didatico_impresso_acessivel = str(linha.iloc[10])
    dict_curso['mat_didatico_impresso_acessivel'] = 0
    if((mat_didatico_impresso_acessivel=="X") or (mat_didatico_impresso_acessivel == "x")):
      dict_curso['mat_didatico_impresso_acessivel'] = 1

    #mat_audio
    mat_audio = str(linha.iloc[11])
    dict_curso['mat_audio'] = 0
    if((mat_audio=="X") or (mat_audio == "x")):
      dict_curso['mat_audio'] = 1

    #mat_impresso_ampliado
    mat_impresso_ampliado = str(linha.iloc[12])
    dict_curso['mat_impresso_ampliado'] = 0
    if((mat_impresso_ampliado=="X") or (mat_impresso_ampliado == "x")):
      dict_curso['mat_impresso_ampliado'] = 1

    #acessibilidade_comunicacao
    acessibilidade_comunicacao = str(linha.iloc[13])
    dict_curso['acessibilidade_comunicacao'] = 0
    if((acessibilidade_comunicacao=="X") or (acessibilidade_comunicacao == "x")):
      dict_curso['acessibilidade_comunicacao'] = 1

    #guia_interprete
    guia_interprete = str(linha.iloc[14])
    dict_curso['guia_interprete'] = 0
    if((guia_interprete=="X") or (guia_interprete == "x")):
      dict_curso['guia_interprete'] = 1


    #libras_ofertada
    libras_ofertada = str(linha.iloc[15])
    dict_curso['libras_ofertada'] = 0
    if((libras_ofertada=="X") or (libras_ofertada == "x")):
      dict_curso['libras_ofertada'] = 1

    #mat_digital_acessivel
    mat_digital_acessivel = str(linha.iloc[16])
    dict_curso['mat_digital_acessivel'] = 0
    if((mat_digital_acessivel=="X") or (mat_digital_acessivel == "x")):
      dict_curso['mat_digital_acessivel'] = 1




    #ao final de cada loop:
    pcd_por_curso[cod_curso] = dict_curso


  return pcd_por_curso


def gerar_txt_migracao_curso(cursos, destino):
  '''recebe uma lista com os dicionários correspondentes a cada curso para ser
  inserido no txt da migração, que será escrito no arquivo de nome <destino>.
  Gera a primeira linha do arquivo de acordo com o leiaute, e depois escreve as próximas
  a partir dos dicionários recebidos

  O cabeçalho do txt de migração deve começar com a linha
  "20|586" , no caso da UFRJ


  chaves do dicionario:
  ['tipo_registro', 'cod_curso', 'aluno_vinculado', 'motivo_sem_aluno',
  'curso_representado_outro_curso', 'convenio',

  'turno_matutino', 'integralizacao_matutino',
  'vagas_novas_matutino', 'vagas_remanescentes_matutino', 'vagas_prog_especiais_matutino',
  'cand_vagas_novas_matutino', 'cand_vagas_remanescentes_matutino', 'cand_vagas_prog_especiais_matutino',
  'turno_vespertino', 'integralizacao_vespertino', 'vagas_novas_vespertino', 'vagas_remanescentes_vespertino'
  'vagas_prog_especiais_vespertino', 'cand_vagas_novas_vespertino', 'cand_vagas_remanescentes_vespertino',
  'cand_vagas_prog_especiais_vespertino', 'turno_noturno', 'integralizacao_noturno', 'vagas_novas_noturno',
  'vagas_remanescentes_noturno', 'vagas_prog_especiais_noturno', 'cand_vagas_novas_noturno',
  'cand_vagas_remanescentes_noturno', 'cand_vagas_prog_especiais_noturno', 'turno_integral',
  'integralizacao_integral', 'vagas_novas_integral', 'vagas_remanescentes_integral', 'vagas_prog_especiais_integral',
  'cand_vagas_novas_integral', 'cand_vagas_remanescentes_integral', 'cand_vagas_prog_especiais_integral',

  'integralizacao_ead', 'vagas_novas_ead', 'vagas_remanescentes_ead', 'vagas_prog_especiais_ead',
  'cand_vagas_novas_ead', 'cand_vagas_remanescentes_ead', 'cand_vagas_prog_especiais_ead',
  'integralizacao_semi', 'vagas_novas_semi', 'vagas_remanescentes_semi', 'vagas_prog_especiais_semi',
  'cand_vagas_novas_semi', 'cand_vagas_remanescentes_semi': '', 'cand_vagas_prog_especiais_semi',


  'acessibilidade', 'mat_braille', 'mat_audio', 'info_acessivel', 'mat_impresso_ampliado',
  'mat_pedagogico_tatil', 'acessibilidade_comunicacao', 'tradutor_libras', 'guia_interprete',
  'mat_didatico_libras', 'libras_ofertada', 'mat_didatico_impresso_acessivel', 'mat_digital_acessivel',
  'oferece_disc_dist', 'percentual_distancia']


  '''

  arq_dest = open(destino, "w+")

  #primeira linha
  arq_dest.write("20|586\n")



  chaves = ['tipo_registro', 'cod_curso', 'aluno_vinculado', 'motivo_sem_aluno',
  'curso_representado_outro_curso', 'convenio',

  'turno_matutino', 'integralizacao_matutino',
  'vagas_novas_matutino', 'vagas_remanescentes_matutino', 'vagas_prog_especiais_matutino',
  'cand_vagas_novas_matutino', 'cand_vagas_remanescentes_matutino', 'cand_vagas_prog_especiais_matutino',
  'turno_vespertino', 'integralizacao_vespertino', 'vagas_novas_vespertino', 'vagas_remanescentes_vespertino',
  'vagas_prog_especiais_vespertino', 'cand_vagas_novas_vespertino', 'cand_vagas_remanescentes_vespertino',
  'cand_vagas_prog_especiais_vespertino', 'turno_noturno', 'integralizacao_noturno', 'vagas_novas_noturno',
  'vagas_remanescentes_noturno', 'vagas_prog_especiais_noturno', 'cand_vagas_novas_noturno',
  'cand_vagas_remanescentes_noturno', 'cand_vagas_prog_especiais_noturno', 'turno_integral',
  'integralizacao_integral', 'vagas_novas_integral', 'vagas_remanescentes_integral', 'vagas_prog_especiais_integral',
  'cand_vagas_novas_integral', 'cand_vagas_remanescentes_integral', 'cand_vagas_prog_especiais_integral',

  'integralizacao_ead', 'vagas_novas_ead', 'vagas_remanescentes_ead', 'vagas_prog_especiais_ead',
  'cand_vagas_novas_ead', 'cand_vagas_remanescentes_ead', 'cand_vagas_prog_especiais_ead',
  'integralizacao_semi', 'vagas_novas_semi', 'vagas_remanescentes_semi', 'vagas_prog_especiais_semi',
  'cand_vagas_novas_semi', 'cand_vagas_remanescentes_semi', 'cand_vagas_prog_especiais_semi',

  'acessibilidade', 'mat_braille', 'mat_audio', 'info_acessivel', 'mat_impresso_ampliado',
  'mat_pedagogico_tatil', 'acessibilidade_comunicacao', 'tradutor_libras', 'guia_interprete',
  'mat_didatico_libras', 'libras_ofertada', 'mat_didatico_impresso_acessivel', 'mat_digital_acessivel',
  'oferece_disc_dist', 'percentual_distancia']


  #print("Dentro do gerar txt")
  print("exemplo de curso: ", cursos[0])

  for curso in cursos:
    linha = ""

    qtd_chaves = len(chaves)

    for i in range(qtd_chaves):
      chave = chaves[i]

      if(i == (qtd_chaves-1)): #se estamos no ultimo
        linha+= str(curso[chave])

      else:

        linha+= str(curso[chave])+"|"


    arq_dest.write(linha+"\n")



  arq_dest.close()

  print("Arquivo de migração pronto: ", destino)

  return




def migracao_curso(nome_origem_curso, nome_origem_pcd, destino, aba_origem_curso="CENSO", aba_origem_pcd = "Acessibilidade_Cursos"):

  #faz um dicionário com as informações a partir da planilha
  #dos cursos, DE ACORDO COM O LEIAUTE. Passa a planhilha de acessibilidade pra função
  # que extrai as infos de acessibilidade para poder gerar o dicionário completo.
  #Usa uma função auxiliar para gerar a planilha de migração
  #IMPORTANTE: SÓ SERVE PARA PLANILHAS COM:
  #           APENAS CURSOS PRESENCIAIS
  #           APENAS CURSOS QUE TIVERAM ALUNO VINCULADO NO ANO DE REFERÊNCIA
  #           FORMATAÇÃO IGUAL À DA PLANILHA DOS CURSOS RECEBIDA PELO NPI NO CASO DOS CURSOS
  #           FORMATAÇÃO IGUAL À DA PLANILHA DE ACESSIBILIDADE RECEBIDA PELO NPI DO CASO DA ACESSIBILIADE
  #NOVAMENTE: RETIRAR DA PLANILHA DOS CURSOS AQUELES QUE NÃO SEJAM 100% PRESENCIAL E TAMBÉM AQUELES
  #QUE NÃO TENHAM ALUNO NO ANO DE REFERÊNCIA



  pcd_por_curso = acessibilidade_cursos(nome_origem_pcd, aba_origem_pcd)



  xls = pd.ExcelFile(nome_origem_curso)
  df1 = pd.read_excel(xls, aba_origem_curso, header=1) #no caso dessa planilha, o cabeçalho ta na segunda linha

  n_linhas, n_col = df1.shape

  cursos = []


  for i in range(n_linhas):
    linha = df1.iloc[i]
    infos_curso = dict()
    #lista_infos_curso = []

    infos_curso['tipo_registro'] = "21"
    infos_curso['cod_curso'] = str(linha.iloc[0])
    cod_curso = str(linha.iloc[0])
    infos_curso['aluno_vinculado'] =  "1"
    infos_curso['motivo_sem_aluno'] =  ""
    infos_curso['curso_representado_outro_curso'] = ""
    infos_curso['convenio'] = "0"

    turno_planilha = (str(linha.iloc[4])).lower() #"matutino", "vespertino", "noturno" ou "integral"

    turnos = ['matutino', 'vespertino', 'noturno', 'integral']

    for turno in turnos:

      if(turno_planilha == turno):
        #se o curso é desse turno, vamos preencher com as infos

        infos_curso['turno_'+turno] = "1"

        integralizacao_turno1 = str(linha.iloc[5]) #converter de (por ex.) 4,5 para '4.5' se não for anos inteiros
        #print("integralização: ", integralizacao_turno1) #DEBUG
        integralizacao_turno = integralizacao_turno1.replace(",", ".")
        #print("integralização correta: ", integralizacao_turno)

        infos_curso['integralizacao_'+turno] = integralizacao_turno

        infos_curso['vagas_novas_'+turno] = str(linha.iloc[6]) #valor em "vagas SISU/ THE/ LIBRAS" na planilha convertido pra string
        infos_curso['vagas_remanescentes_'+turno] = str(linha.iloc[7]) #"vagas remanescentes" na planilha

        vagas_prog_especiais = str(linha.iloc[8])
        #print("vagas prog especiais: ", vagas_prog_especiais)
        if(( (vagas_prog_especiais.lower()) == 'nan') or vagas_prog_especiais == ""):
          vagas_prog_especiais = "0"


        infos_curso['vagas_prog_especiais_'+ turno] = vagas_prog_especiais

        infos_curso['cand_vagas_novas_'+turno] = str(linha.iloc[9])
        infos_curso['cand_vagas_remanescentes_'+turno] = str(linha.iloc[10])

        cand_vagas_prog_especiais = str(linha.iloc[11])
        if(( (cand_vagas_prog_especiais.lower()) == 'nan') or cand_vagas_prog_especiais == ""):
          cand_vagas_prog_especiais = "0"

        infos_curso['cand_vagas_prog_especiais_'+ turno] = cand_vagas_prog_especiais

      else: #vai ficar em branco em relação a esse turno
        infos_curso['turno_'+turno] = "0"

        infos_curso['integralizacao_'+turno] = ""

        infos_curso['vagas_novas_'+turno] = ""
        infos_curso['vagas_remanescentes_'+turno] = ""

        infos_curso['vagas_prog_especiais_'+ turno] = ""

        infos_curso['cand_vagas_novas_'+turno] = ""
        infos_curso['cand_vagas_remanescentes_'+turno] = ""

        infos_curso['cand_vagas_prog_especiais_'+ turno] = ""


    #Ead (em branco)
    infos_curso['integralizacao_ead'] = ""

    infos_curso['vagas_novas_ead'] = ""
    infos_curso['vagas_remanescentes_ead'] = ""

    infos_curso['vagas_prog_especiais_ead'] = ""

    infos_curso['cand_vagas_novas_ead'] = ""
    infos_curso['cand_vagas_remanescentes_ead'] = ""

    infos_curso['cand_vagas_prog_especiais_ead'] = ""


    #semipresencial (em branco)

    infos_curso['integralizacao_semi'] = ""

    infos_curso['vagas_novas_semi'] = ""
    infos_curso['vagas_remanescentes_semi'] = ""

    infos_curso['vagas_prog_especiais_semi'] = ""

    infos_curso['cand_vagas_novas_semi'] = ""
    infos_curso['cand_vagas_remanescentes_semi'] = ""

    infos_curso['cand_vagas_prog_especiais_semi'] = ""



    '''acessibilidade : "1" se o codigo do curso constar no dicionario retornado
                  pela função que informa acessibilidade (acessibilidade_cursos) E se o valor "1" constar
                  em algum dos valores do dicionário correspondente ao curso; "0" caso contrário

      (NESSA ORDEM:) mat_braille, mat_audio, info_acessivel, mat_impresso_ampliado, mat_pedagogico_tatil,
      acessibilidade_comunicacao, tradutor_libras, guia_interprete, mat_didatico_libras, libras_ofertada,
      mat_didatico_impresso_acessivel, mat_digital_acessivel:  preencher de acordo com o
      dicionario de acessibilidade'''

    pcd_em_branco = 1
    if(cod_curso in pcd_por_curso.keys()):
      dict_curso = pcd_por_curso[cod_curso]

      if(1 in dict_curso.values()): #significa que o curso de fato oferece acessibilidade

        pcd_em_branco = 0

        infos_curso['acessibilidade'] = "1"

        infos_curso['mat_braille'] = str(dict_curso['mat_braille'])

        infos_curso['mat_audio'] = str(dict_curso['mat_audio'])


        infos_curso['info_acessivel'] = str(dict_curso['info_acessivel'])

        infos_curso['mat_impresso_ampliado'] = str(dict_curso['mat_impresso_ampliado'])

        infos_curso['mat_pedagogico_tatil'] = str(dict_curso['mat_pedagogico_tatil'])

        infos_curso['acessibilidade_comunicacao'] = str(dict_curso['acessibilidade_comunicacao'])

        infos_curso['tradutor_libras'] = str(dict_curso['tradutor_libras'])

        infos_curso['guia_interprete'] = str(dict_curso['guia_interprete'])

        infos_curso['mat_didatico_libras'] = str(dict_curso['mat_didatico_libras'])

        infos_curso['libras_ofertada'] = str(dict_curso['libras_ofertada'])

        infos_curso['mat_didatico_impresso_acessivel'] = str(dict_curso['mat_didatico_impresso_acessivel'])

        infos_curso['mat_digital_acessivel'] = str(dict_curso['mat_digital_acessivel'])


    if(pcd_em_branco == 1): #idealmente não é pra acontecer
      infos_curso['acessibilidade'] = "0"

      infos_curso['mat_braille'] = ""

      infos_curso['mat_audio'] = ""

      infos_curso['info_acessivel'] = ""

      infos_curso['mat_impresso_ampliado'] = ""

      infos_curso['mat_pedagogico_tatil'] = ""

      infos_curso['acessibilidade_comunicacao'] = ""

      infos_curso['tradutor_libras'] = ""

      infos_curso['guia_interprete'] = ""

      infos_curso['mat_didatico_libras'] = ""

      infos_curso['libras_ofertada'] = ""

      infos_curso['mat_didatico_impresso_acessivel'] = ""

      infos_curso['mat_digital_acessivel'] = ""



    #falta os dois ultimos campos
    '''oferece_disc_dist: "0" sempre

    percentual_distancia: "" sempre'''

    infos_curso['oferece_disc_dist'] = '0'
    infos_curso['percentual_distancia'] = ""



    #no final de cada loop
    cursos.append(infos_curso)




  #print(cursos[0])


  #passar pro arquivo
  gerar_txt_migracao_curso(cursos, destino)
  return




if __name__ == "__main__":


  nome_origem_curso = input("Insira o nome da planilha dos cursos (exemplo: cursos.xlsx)\n")
  aba_origem_curso = input("Insira o nome da aba dessa planilha com as informações dos cursos (exemplo: Cursos)\n")
  nome_origem_pcd = input("Insira o nome da planilha com as informações de acessibilidade por curso (exemplo: acessibilidade.xlsx)\n")
  aba_origem_pcd = input("Insira o nome da aba dessa planilha com as informações dos cursos (exemplo: Acessibilidade)\n")

  destino = "migracao_curso.txt"

  migracao_curso(nome_origem_curso, nome_origem_pcd, destino, aba_origem_curso, aba_origem_pcd)

