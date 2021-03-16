#-----------------------------------------------------------------#
#-----------------------------------------------------------------#
#----------------- FORMATO AOM - 502 -----------------------------#
#-----------------------------------------------------------------#
#-----------------------------------------------------------------#


######################################################################
#                                                                    #
#       la información cargada en este archivo hace parte de la      #
#       base de datos que se debe construir para cumplir con el      #
#       reporte de información de la Circular CREG 114 de 2019       #
#       en lo referente a diligenciar el archivo "FORMATO AOM-502"   # 
#                                                                    #
######################################################################

#    La base de datos consultada pertenece al descargue que se hace del sustema de información de TGI
#    Contenido en la herramienta SAP

getwd
#------ set path in TGI PC ----------#
setwd("E://OneDrive - TRANSPORTADORA DE GAS INTERNACIONAL S.A. ESP//aorozco//Mis documentos//2020//AOM//Circular-CREG-114-2019//TGI-info_Areas")

#------ set path in personal PC ----------#
setwd("C://Users//aleja//Downloads//tgi//AOM")

#Installing aditional packages
install.packages("reshape2")
install.packages("ggthemes")
install.packages("extrafont")
install.packages("ggplot2")
install.packages("dplyr")

## para leer excel
install.packages("readxl")
install.packages("XLConnect")
install.packages("XLConnectJars")
install.packages("xlsx")

#PLOTLY
install.packages("plotly")

#Calling library of packages
library(reshape2)
library(ggthemes)
library(ggplot2)
library(extrafont)

#PLOTLY
library(plotly)

#other library
library(scales)
library(dplyr)

#para excel
library(readxl)
library(XLConnect)

# Read all worksheets in an Excel workbook into an R list with data.frames

#Cargando datos de la base ya creada con anterioridad - este es un comando nuevo
#Generacion_2015_2018<-read.csv("Generacion_2015_2018.csv",header=T,sep=",")

#convertir datos de fecha en formato "date"
#Generacion_2015_2018$Fecha<- as.Date(Generacion_2015_2018$Fecha, format='%Y-%m-%d')

  ########################################
  ######------para borrar todo------######
  
  rm(list=ls())
  
  ########################################
  ########################################
  
  #para leer todas las hojas de un arvhivo de excel
  library(readxl)    
  read_excel_allsheets <- function(filename, tibble = FALSE) {
  # I prefer straight data.frames
  # but if you like tidyverse tibbles (the default with read_excel)
  # then just pass tibble = TRUE
  sheets <- readxl::excel_sheets(filename)
  x <- lapply(sheets, function(X) readxl::read_excel(filename, sheet = X))
  if(!tibble) x <- lapply(x, as.data.frame)
  names(x) <- sheets
  x
  }

  mysheets <- read_excel_allsheets("AOM CONCEPTOS A EXCLUIR CON CAPEX 25 02-2020 ASV.xlsx")

  #Leer archivo de excel  
  matrizConciliacion=readWorksheetFromFile("MatrizDeConciliacion.xlsx", sheet = 1)
  
  options(scipen=999) #Para que no tenga notacion cientifica 
  
  #create an empty data frame
  df <- data.frame(Characters=character(),
                   Ints=integer(),
                   Ints=integer(),
                   Ints=integer(),
                   Ints=integer(),
                   Characters=character(),
                   Characters=character(),
                   Doubles=double(),
                   stringsAsFactors=FALSE)
  colnames(df) <- c("Tramo", "Código_SSPD_PUC", "Código_CGN_PUC", "Código_CREG_ICR", "Código_TGI_contabilidad","Nombre_Concepto_CREG_IGR/Cuenta_PUC", "Nombre_Concepto_TGI_Contabilidad", "Valor")

  #entrega el nombre de cada una de las hojas a revisar dentro del contenedor "mysheets"
  #names(mysheets[1])
  #para revisar la base de datos de ballena, despues se debe eliminar
  bb=mysheets$`Ballena - Barranca`
  bs=mysheets$`Barranca - Sebastopol`
  r=1
  k=1
  l=1
  z=0
  # se recorre la "contenedora" de base de datos con la información de cada tramo
  for (j in 20:21)#(length(mysheets)-7)
  {
    #se recorre cada una de las filas de la matriz de conciliación par poder comparar el numero de cuenta PUC con la información de TGI
    for (i in 4:nrow(matrizConciliacion)) #23
    {
      #se define el tramo en el cual inicia la comparación
      tramo = mysheets[[j]]
      # se recorren las filas del tramos para posteriormente hacer la comparación de cuentas con "matrizConciliacion" 
      for (k in 2:nrow(tramo))
      {
        #extrae los numeros que quiero comparar, en este caso se extraen 6
        compMconcil6=substr(tramo[k,1],1,6)
        #extrae los numeros que quiero comparar, en este caso se extraen 4
        compMconcil4=substr(tramo[k,1],1,4)
        #extrae los numeros que quiero comparar, en este caso se extraen 2
        compMconcil2=substr(tramo[k,1],1,2)
          
        #readline(prompt="Press [enter] to continue")
        if (is.na(compMconcil6)!=TRUE)#|is.na(compMconcil4)!=TRUE)
        {
          if (is.na(matrizConciliacion[i,1])!=TRUE|is.na(matrizConciliacion[i,2])!=TRUE)
          {
            #para comparar las cuentas de "matrizConciliacion" con "tramo" se saca el numeri de caracteres con "nchar"
            if (nchar(matrizConciliacion[i,1])==6|nchar(matrizConciliacion[i,2])==6)
            {
              if (matrizConciliacion[i,1]==compMconcil6|matrizConciliacion[i,2]==compMconcil6)
              {
                df[l,1]= names(mysheets[j])
                df[l,2]= matrizConciliacion[i,1]
                df[l,3]= matrizConciliacion[i,2]
                df[l,4]= matrizConciliacion[i,3]
                df[l,5]= tramo[k,1]
                df[l,6]= matrizConciliacion[i,4]
                df[l,7]= tramo[k,2]
                df[l,8]= tramo[k,10]
                print(l)
                l=l+1
              }
            }
            if (nchar(matrizConciliacion[i,1])==4|nchar(matrizConciliacion[i,2])==4)
            {
              if (matrizConciliacion[i,1]==compMconcil4|matrizConciliacion[i,2]==compMconcil4)
              {
                df[l,1]= names(mysheets[j])
                df[l,2]= matrizConciliacion[i,1]
                df[l,3]= matrizConciliacion[i,2]
                df[l,4]= matrizConciliacion[i,3]
                df[l,5]= tramo[k,1]
                df[l,6]= matrizConciliacion[i,4]
                df[l,7]= tramo[k,2]
                df[l,8]= tramo[k,10]
                print(l)
                l=l+1
              }
            }
            ####aqui ingresa la comparción con las cuentas de dos digitos que se encuentra en
            #### discusión si se incluyen y se dejan unicamente las de 4 digitos
            if (nchar(matrizConciliacion[i,1])==2|nchar(matrizConciliacion[i,2])==2)
            {
              if (matrizConciliacion[i,1]==compMconcil2|matrizConciliacion[i,2]==compMconcil2)
              {
                df[l,1]= names(mysheets[j])
                df[l,2]= matrizConciliacion[i,1]
                df[l,3]= matrizConciliacion[i,2]
                df[l,4]= matrizConciliacion[i,3]
                df[l,5]= tramo[k,1]
                df[l,6]= matrizConciliacion[i,4]
                df[l,7]= tramo[k,2]
                df[l,8]= tramo[k,10]
                print(l)
                l=l+1
              }
            }
          }
        }
        #z=z+1
        #readline(prompt="Press [enter] to continue")
      }  
      #r=r+1
      #readline(prompt="Press [enter] to continue")
    }
  }
  df[,8]=as.numeric(df[,8])
  ###----------------------------------------------------------------------------------------------------###
  ###---- Se crea el dataframe que va a contener las cuentas que no estan repetidas para cada tramo -----###
  ###----------------------------------------------------------------------------------------------------###
  
  
  #create an empty data frame
  df3 <- data.frame(Characters=character(),
                   Ints=integer(),
                   Ints=integer(),
                   Ints=integer(),
                   Ints=integer(),
                   Characters=character(),
                   Characters=character(),
                   Doubles=double(),
                   stringsAsFactors=FALSE)
  colnames(df3) <- c("Tramo", "Código_SSPD_PUC", "Código_CGN_PUC", "Código_CREG_ICR", "Código_TGI_contabilidad","Nombre_Concepto_CREG_IGR/Cuenta_PUC", "Nombre_Concepto_TGI_Contabilidad", "Valor")
  
  # se recorre la "contenedora" de base de datos con la información de cada tramo
  for (j in 20:21)#(length(mysheets)-7)
  {
    #
    df2=subset(df,Tramo==names(mysheets[j]))
    #Se eliminan los valores repetidos
    df2=distinct(df2,df2[,5], .keep_all= TRUE)
    #Se concatenan las bases de datos sin cuentas repetidas para dejarlo en una unica matriz
    df3=rbind(df3,df2)
  }
  
  
  ###-------------------------------------------------------------------------------------------------------------###
  ###--- final de la creación dataframe que va a contener las cuentas que no estan repetidas para cada tramo  ----###
  ###-------------------------------------------------------------------------------------------------------------###
  
  #myworkbook
  #matriz_formatoAOM_502
  library("xlsx")
  write.xlsx(df3, file = "myworkbook.xlsx",
             sheetName = "df3", append = FALSE)
  
  # calculos de tabla dinamica dcast o melt
  #+df2$Código_CREG_ICR
  
  total=dcast(df3,df3$Código_CREG_ICR~df3$Tramo,value.var = "Valor",sum)
  probando=subset(df3,df3$Código_CREG_ICR=="01010201")

