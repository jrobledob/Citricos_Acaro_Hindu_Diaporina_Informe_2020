library(dplyr)
library(ggplot2)
library(officer)
library(rvg)
library(ggplotify)
library(tidyverse)
library(lubridate)

###Ácaro Hindú
#Reading the data
acaro_BD<- read.csv2("./Data/BD_acaro_2020.csv", header = T, encoding = "UTF-8")
#Filtering acaro_BD according to date. All data obtained from october 8 2019 to present
#and "evaluado"== B are selected
acaro_2020<- acaro_BD %>%
  filter(MUESTREO>="M69" & EVALUADO!=c("B1","B2"))
##Figure by citrus type
#Set "Fecha" up
acaro_2020$Fecha<- as.Date(paste(acaro_2020$AÑO,acaro_2020$MES,acaro_2020$DIA, sep="-"))

##EGGS
#Get mean values per citrus type and sample (date). Contongency table
acaro_2020$HUEVO.ACARO.HINDU<- as.numeric(as.character(acaro_2020$HUEVO.ACARO.HINDU))
mean_by_citrus_type<- aggregate(acaro_2020$HUEVO.ACARO.HINDU, list(acaro_2020$CULTIVAR, acaro_2020$Fecha), mean, na.rm=T)
#set names of mean_by_citrus_type
names(mean_by_citrus_type)<- c("Variedad", "Fecha", "Número de Huevos Promedio por Hoja")
figure_mean_by_citrus_type<- ggplot(mean_by_citrus_type, aes(x=Fecha, y=`Número de Huevos Promedio por Hoja`, group=Variedad))+
  geom_line(aes(linetype=Variedad))+
  geom_point(aes(shape=Variedad))+
  theme_classic()+
  theme(axis.text.x = element_text(angle = 45))
#save as PPT
p_dml <-dml(ggobj = figure_mean_by_citrus_type)
read_pptx() %>%
  add_slide() %>%
  ph_with(p_dml, ph_location()) %>%
  base::print("1. figure_mean_by_citrus_type.pptx")


#Get mean values per rootstock and sample (date). Contongency table
mean_by_citrus_type<- aggregate(acaro_2020$HUEVO.ACARO.HINDU, list(acaro_2020$PORTAINJERTO, acaro_2020$Fecha), mean, na.rm=T)
#set names of mean_by_citrus_type
names(mean_by_citrus_type)<- c("Porta Injerto", "Fecha", "Número de Huevos Promedio por Hoja")
figure_mean_by_citrus_type<- ggplot(mean_by_citrus_type, aes(x=Fecha, y=`Número de Huevos Promedio por Hoja`, group=`Porta Injerto`))+
  geom_line(aes(linetype=`Porta Injerto`))+
  geom_point(aes(shape=`Porta Injerto`))+
  theme_classic()+
  theme(axis.text.x = element_text(angle = 45))
#save as PPT
p_dml <-dml(ggobj = figure_mean_by_citrus_type)
read_pptx() %>%
  add_slide() %>%
  ph_with(p_dml, ph_location()) %>%
  base::print("2. figure_mean_by_rootstock.pptx")


#Get mean values per variety per citrus type and sample (date). Contongency table
mean_by_citrus_type<- acaro_2020%>%
  filter(CULTIVAR=="Naranja")
mean_by_citrus_type<- aggregate(mean_by_citrus_type$HUEVO.ACARO.HINDU, list(mean_by_citrus_type$VARIEDAD, mean_by_citrus_type$Fecha), mean, na.rm=T)
#set names of mean_by_citrus_type
names(mean_by_citrus_type)<- c("Variedad", "Fecha", "Número de Huevos Promedio por Hoja")
figure_mean_by_citrus_type<- ggplot(mean_by_citrus_type, aes(x=Fecha, y=`Número de Huevos Promedio por Hoja`, group=Variedad))+
  geom_line(aes(color=Variedad))+
  geom_point(aes(color=Variedad))+
  theme_classic()+
  theme(axis.text.x = element_text(angle = 45))
#save as PPT
p_dml <-dml(ggobj = figure_mean_by_citrus_type)
read_pptx() %>%
  add_slide() %>%
  ph_with(p_dml, ph_location()) %>%
  base::print("3. figure_mean_by_variety_Oranges.pptx")

#Get mean values per variety per citrus type and sample (date). Contongency table
mean_by_citrus_type<- acaro_2020%>%
  filter(CULTIVAR=="Mandarina")
mean_by_citrus_type<- aggregate(mean_by_citrus_type$HUEVO.ACARO.HINDU, list(mean_by_citrus_type$VARIEDAD, mean_by_citrus_type$Fecha), mean, na.rm=T)
#set names of mean_by_citrus_type
names(mean_by_citrus_type)<- c("Variedad", "Fecha", "Número de Huevos Promedio por Hoja")
figure_mean_by_citrus_type<- ggplot(mean_by_citrus_type, aes(x=Fecha, y=`Número de Huevos Promedio por Hoja`, group=Variedad))+
  geom_line(aes(color=Variedad))+
  geom_point(aes(color=Variedad))+
  theme_classic()+
  theme(axis.text.x = element_text(angle = 45))
#save as PPT
p_dml <-dml(ggobj = figure_mean_by_citrus_type)
read_pptx() %>%
  add_slide() %>%
  ph_with(p_dml, ph_location()) %>%
  base::print("4. figure_mean_by_variety_Tangerines.pptx")

#Get mean values per variety per citrus type and sample (date). Contongency table
mean_by_citrus_type<- acaro_2020%>%
  filter(CULTIVAR=="Limón")
mean_by_citrus_type<- aggregate(mean_by_citrus_type$HUEVO.ACARO.HINDU, list(mean_by_citrus_type$VARIEDAD, mean_by_citrus_type$Fecha), mean, na.rm=T)
#set names of mean_by_citrus_type
names(mean_by_citrus_type)<- c("Variedad", "Fecha", "Número de Huevos Promedio por Hoja")
figure_mean_by_citrus_type<- ggplot(mean_by_citrus_type, aes(x=Fecha, y=`Número de Huevos Promedio por Hoja`, group=Variedad))+
  geom_line(aes(color=Variedad))+
  geom_point(aes(color=Variedad))+
  theme_classic()+
  theme(axis.text.x = element_text(angle = 45))
#save as PPT
p_dml <-dml(ggobj = figure_mean_by_citrus_type)
read_pptx() %>%
  add_slide() %>%
  ph_with(p_dml, ph_location()) %>%
  base::print("5. figure_mean_by_variety_Lemons.pptx")


##Adults
#Get mean values per citrus type and sample (date). Contongency table
acaro_2020$ADULTO.ACARO.HINDU<- as.numeric(as.character(acaro_2020$ADULTO.ACARO.HINDU))
mean_by_citrus_type<- aggregate(acaro_2020$ADULTO.ACARO.HINDU, list(acaro_2020$CULTIVAR, acaro_2020$Fecha), mean, na.rm=T)
#set names of mean_by_citrus_type
names(mean_by_citrus_type)<- c("Variedad", "Fecha", "Número Promedio de Adultos por Hoja")
figure_mean_by_citrus_type<- ggplot(mean_by_citrus_type, aes(x=Fecha, y=`Número Promedio de Adultos por Hoja`, group=Variedad))+
  geom_line(aes(linetype=Variedad))+
  geom_point(aes(shape=Variedad))+
  theme_classic()+
  theme(axis.text.x = element_text(angle = 45))
#save as PPT
p_dml <-dml(ggobj = figure_mean_by_citrus_type)
read_pptx() %>%
  add_slide() %>%
  ph_with(p_dml, ph_location()) %>%
  base::print("6. figure_mean_by_citrus_type_Adults.pptx")


#Get mean values per rootstock and sample (date). Contongency table
mean_by_citrus_type<- aggregate(acaro_2020$ADULTO.ACARO.HINDU, list(acaro_2020$PORTAINJERTO, acaro_2020$Fecha), mean, na.rm=T)
#set names of mean_by_citrus_type
names(mean_by_citrus_type)<- c("Porta Injerto", "Fecha", "Número Promedio de Adultos por Hoja")
figure_mean_by_citrus_type<- ggplot(mean_by_citrus_type, aes(x=Fecha, y=`Número Promedio de Adultos por Hoja`, group=`Porta Injerto`))+
  geom_line(aes(linetype=`Porta Injerto`))+
  geom_point(aes(shape=`Porta Injerto`))+
  theme_classic()+
  theme(axis.text.x = element_text(angle = 45))
#save as PPT
p_dml <-dml(ggobj = figure_mean_by_citrus_type)
read_pptx() %>%
  add_slide() %>%
  ph_with(p_dml, ph_location()) %>%
  base::print("7. figure_mean_by_rootstock_Adults.pptx")


#Get mean values per variety per citrus type and sample (date). Contongency table
mean_by_citrus_type<- acaro_2020%>%
  filter(CULTIVAR=="Naranja")
mean_by_citrus_type<- aggregate(mean_by_citrus_type$ADULTO.ACARO.HINDU, list(mean_by_citrus_type$VARIEDAD, mean_by_citrus_type$Fecha), mean, na.rm=T)
#set names of mean_by_citrus_type
names(mean_by_citrus_type)<- c("Variedad", "Fecha", "Número Promedio de Adultos por Hoja")
figure_mean_by_citrus_type<- ggplot(mean_by_citrus_type, aes(x=Fecha, y=`Número Promedio de Adultos por Hoja`, group=Variedad))+
  geom_line(aes(color=Variedad))+
  geom_point(aes(color=Variedad))+
  theme_classic()+
  theme(axis.text.x = element_text(angle = 45))
#save as PPT
p_dml <-dml(ggobj = figure_mean_by_citrus_type)
read_pptx() %>%
  add_slide() %>%
  ph_with(p_dml, ph_location()) %>%
  base::print("8. figure_mean_by_variety_Oranges_Adults.pptx")

#Get mean values per variety per citrus type and sample (date). Contongency table
mean_by_citrus_type<- acaro_2020%>%
  filter(CULTIVAR=="Mandarina")
mean_by_citrus_type<- aggregate(mean_by_citrus_type$ADULTO.ACARO.HINDU, list(mean_by_citrus_type$VARIEDAD, mean_by_citrus_type$Fecha), mean, na.rm=T)
#set names of mean_by_citrus_type
names(mean_by_citrus_type)<- c("Variedad", "Fecha", "Número Promedio de Adultos por Hoja")
figure_mean_by_citrus_type<- ggplot(mean_by_citrus_type, aes(x=Fecha, y=`Número Promedio de Adultos por Hoja`, group=Variedad))+
  geom_line(aes(color=Variedad))+
  geom_point(aes(color=Variedad))+
  theme_classic()+
  theme(axis.text.x = element_text(angle = 45))
#save as PPT
p_dml <-dml(ggobj = figure_mean_by_citrus_type)
read_pptx() %>%
  add_slide() %>%
  ph_with(p_dml, ph_location()) %>%
  base::print("9. figure_mean_by_variety_Tangerines_Adults.pptx")

#Get mean values per variety per citrus type and sample (date). Contongency table
mean_by_citrus_type<- acaro_2020%>%
  filter(CULTIVAR=="Limón")
mean_by_citrus_type<- aggregate(mean_by_citrus_type$ADULTO.ACARO.HINDU, list(mean_by_citrus_type$VARIEDAD, mean_by_citrus_type$Fecha), mean, na.rm=T)
#set names of mean_by_citrus_type
names(mean_by_citrus_type)<- c("Variedad", "Fecha", "Número Promedio de Adultos por Hoja")
figure_mean_by_citrus_type<- ggplot(mean_by_citrus_type, aes(x=Fecha, y=`Número Promedio de Adultos por Hoja`, group=Variedad))+
  geom_line(aes(color=Variedad))+
  geom_point(aes(color=Variedad))+
  theme_classic()+
  theme(axis.text.x = element_text(angle = 45))
#save as PPT
p_dml <-dml(ggobj = figure_mean_by_citrus_type)
read_pptx() %>%
  add_slide() %>%
  ph_with(p_dml, ph_location()) %>%
  base::print("10. figure_mean_by_variety_Lemons_Adults.pptx")







##Otros Ácaros
#Get mean values per citrus type and sample (date). Contongency table
acaro_2020$OTROS.ACAROS<- as.numeric(as.character(acaro_2020$OTROS.ACAROS))
mean_by_citrus_type<- as.data.frame(table(acaro_2020$OTROS.ACAROS, acaro_2020$CULTIVAR, acaro_2020$Fecha))
#set names of mean_by_citrus_type
names(mean_by_citrus_type)<- c("Severidad", "Cultivar", "Fecha", "Frecuencia")
mean_by_citrus_type$Fecha<- as.Date(mean_by_citrus_type$Fecha)

mean_by_citrus_type2 <- mean_by_citrus_type %>% 
  group_by(FECHA = floor_date(Fecha, unit = "month"))%>%
  summarize(Frecuencia = sum(Frecuencia),  Severidad= Severidad, Cultivar=
              Cultivar)




#Figure
figure_mean_by_citrus_type<- ggplot(mean_by_citrus_type2, aes(x=Fecha, y=Frecuencia, group=Cultivar, fill=Cultivar, alpha=Severidad)) + 
      geom_bar(stat="identity",position="dodge", colour="black") +
      scale_alpha_manual(values=c(0,0.25, 0.5, 0.75, 1))
#save as PPT
p_dml <-dml(ggobj = figure_mean_by_citrus_type)
read_pptx() %>%
  add_slide() %>%
  ph_with(p_dml, ph_location()) %>%
  base::print("11. figure_mean_by_variety_citrus_types_Otros_Ácaros.pptx")

