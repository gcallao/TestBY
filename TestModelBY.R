library("readxl")
library("dplyr")
library("forecast")

# 1) Carga de Data--------

d01 <- read_excel("C:/Users/gcall/Downloads/Test Fcst/Historico APO CEDI SKU/PE_01.2020_30.2020.xlsx")
d02 <- read_excel("C:/Users/gcall/Downloads/Test Fcst/Historico APO CEDI SKU/PE_31.2020_40.2020.xlsx")
d03 <- read_excel("C:/Users/gcall/Downloads/Test Fcst/Historico APO CEDI SKU/PE_41.2020_53.2020.xlsx")
d04 <- read_excel("C:/Users/gcall/Downloads/Test Fcst/Historico APO CEDI SKU/PE_01.2021_20.2021.xlsx")
d05 <- read_excel("C:/Users/gcall/Downloads/Test Fcst/Historico APO CEDI SKU/PE_21.2021_40.2021.xlsx")
d06 <- read_excel("C:/Users/gcall/Downloads/Test Fcst/Historico APO CEDI SKU/PE_41.2021_52.2021.xlsx")
d07 <- read_excel("C:/Users/gcall/Downloads/Test Fcst/Historico APO CEDI SKU/PE_01.2022_20.2022.xlsx")
d08 <- read_excel("C:/Users/gcall/Downloads/Test Fcst/Historico APO CEDI SKU/PE_21.2022_22.2022.xlsx")

productos <- read_excel("C:/Users/gcall/Downloads/Test Fcst/productos.xlsx")
cedis <- read_excel("C:/Users/gcall/Downloads/Test Fcst/cedis.xlsx")

data <- union_all(d01, d02)
data <- union_all(data, d03)
data <- union_all(data, d04)
data <- union_all(data, d05)
data <- union_all(data, d06)
data <- union_all(data, d07)
data <- union_all(data, d08)
rm(d01); rm(d02); rm(d03); rm(d04); rm(d05); rm(d06); rm(d07); rm(d08)

# 2) Preprocesamiento de Data--------

data <- merge(data, productos, by = "ZAC_PROD")
data <- merge(data, cedis, by = "CEDI")

# Formación de la clave
data$combo_ag <- paste(data$SUBREGION_CDA, data$agrupacion, data$segmento_agrupado, data$marca, data$tamaño, data$retornable)

# Gráfico exploratorio
resumen <- aggregate(data[,6], by = list(CALWEEK = data$CALWEEK), FUN = sum, na.rm = TRUE)
colnames(resumen) <- c("CALWEEK", "HVTA")

plot(resumen$HVTA, type = "l")


# Imputar valores 0 del ratio de venta ajustada
data$ZAC_HVNA[data$CALWEEK>=202214] <- data$ZAC_HVTA[data$CALWEEK>=202214]

# Generar data agregada a nivel nacional
data_ag <- aggregate(data[,c(5,6)], by = list(SUBREGION_CDA = data$SUBREGION_CDA, agrupacion = data$agrupacion,
                                              segmento_agrupado = data$segmento_agrupado, marca = data$marca,
                                              tamaño = data$tamaño, retornable = data$retornable, CALWEEK = data$CALWEEK)
                     , FUN = sum, na.rm = TRUE) 

data_ag$id <- paste(data_ag$SUBREGION_CDA, data_ag$agrupacion, data_ag$segmento_agrupado, 
                    data_ag$marca, data_ag$tamaño, data_ag$retornable)

# Lista de combinaciones a nivel detallado (1-4-1)
combo_total <- aggregate(data[,6], by = list(CEDI = data$CEDI, ZAC_PROD = data$ZAC_PROD, locacion = data$LOCACION, producto = data$producto)
                         , FUN = sum, na.rm = TRUE) 

colnames(combo_total)[5] <- "HVTA"

combo_total <- combo_total[order(combo_total$HVTA, decreasing = TRUE),]
combo_total <- combo_total[combo_total$HVTA>0,]

combo_total$id <- paste(combo_total$CEDI, combo_total$ZAC_PROD)
row.names(combo_total) <- NULL

# Lista de combinaciones a nivel agregado (2-4-3)
combo_ag <- aggregate(data[,6], by = list(SUBREGION_CDA = data$SUBREGION_CDA, agrupacion = data$agrupacion,
                                           segmento_agrupado = data$segmento_agrupado, marca = data$marca,
                                           tamaño = data$tamaño, retornable = data$retornable)
                      , FUN = sum, na.rm = TRUE) 

colnames(combo_ag)[7] <- "HVTA"

combo_ag <- combo_ag[order(combo_ag$HVTA, decreasing = TRUE),]
combo_ag <- combo_ag[combo_ag$HVTA>0,]

# Formación de la clave - nivel agregado
combo_ag$id <- paste(combo_ag$SUBREGION_CDA, combo_ag$agrupacion, combo_ag$segmento_agrupado, combo_ag$marca, combo_ag$tamaño, combo_ag$retornable)
row.names(combo_ag) <- NULL


# 3) Modelamiento Estadístico----

combo <- combo_ag$id[8]

"SUBREGION_CDA  | AGRUPACIÓN  | SEGMENTO_AGRUPADO   | MARCA     | TAMAÑO      | RETORNABLE    |" 
"--------------------------------------------------------------------------------------------------"
"OLS LIMA       | GASEOSAS    | MIXTOS              | MIXTOS    | FAMILIAR    | NO RETORNABLE |"

# 3.1) Modelo Agregado

x <- subset(data_ag, id == combo)
x$predy_cf_ag <- 0


x.ts <- ts(subset(x, CALWEEK <= 202212)$ZAC_HVNA, frequency = 52)
n.ts <- length(x.ts)
n.x <- nrow(x)

x.nn <- stlf(x.ts)

x.nn.pred <- forecast(x.nn, h = 10)
x$predy_cf_ag[1:(n.x-10)] <- x.nn.pred$fitted
x$predy_cf_ag[(n.x-10+1):n.x] <- x.nn.pred$mean
n <- x.nn.pred

x$error_abs_cf <- abs(x$ZAC_HVTA-x$predy_cf_ag)
x$acc_cf <- 1-x$error_abs_cf/x$ZAC_HVTA

x$combo_ag_sem <- paste(x$id, x$CALWEEK)

par(mfrow=c(2,1))
plot(x$ZAC_HVTA, main = "Historia de Venta", type = "l", cex.main=1.5)
plot(n, main = "Pronóstico en Cajas Físicas", cex.main=1.5)


# 3.2) Modelo Detallado

# Generación data detallada
df <- data[data$combo_ag==combo,] #Filtrando por combo del nivel agrupado
df$id <- paste(df$CEDI, df$ZAC_PROD)

# Listado de combinaciones detalladas
combo_id <- aggregate(df[,6], by = list(CEDI = df$CEDI, ZAC_PROD = df$ZAC_PROD, locacion = df$LOCACION, producto = df$producto)
                      , FUN = sum, na.rm = TRUE) 

colnames(combo_id)[5] <- "HVTA"

combo_id <- combo_id[order(combo_id$HVTA, decreasing = TRUE),]

combo_id$id <- paste(combo_id$CEDI, combo_id$ZAC_PROD)
row.names(combo_id) <- NULL

df$predy_cf_det <- 0
df$error_abs_cf <- 0
df$acc_cf <- 0

bd <- df[-(1:nrow(df)),]

for (i in 1:nrow(combo_id)){
  
  print(i)
  
  c_id <- combo_id$id[i]
  
  y <- subset(df, id == c_id)
  y <- y[order(y$CALWEEK, decreasing = FALSE),]
  
  y.ts <- ts(subset(y, CALWEEK <= 202212)$ZAC_HVNA, frequency = 52)
  n.ts <- length(y.ts)
  n.y <- nrow(y)
  
  y.nn <- stlf(y.ts)
  
  y.nn.pred <- forecast(y.nn, h = 10)
  y$predy_cf_det[1:(n.y-10)] <- y.nn.pred$fitted
  y$predy_cf_det[(n.y-10+1):n.y] <- y.nn.pred$mean
  n <- y.nn.pred
  
  #Limpiar pronósticos menores a 0
  y$predy_cf_det[y$CALWEEK >= 202213 & y$predy_cf < 0] <- 0 
  
  #Retirar pronóstico SKUS descontinuados
  y$predy_cf_det[y$ZAC_PROD ==  "PE_251385" | y$ZAC_PROD ==  "PE_254850" |  y$ZAC_PROD ==  "PE_256489" | 
                y$ZAC_PROD ==  "PE_256490" | y$ZAC_PROD ==  "PE_256491" | y$ZAC_PROD ==  "PE_256492"] <- 0
  y$error_abs_cf <- abs(y$ZAC_HVTA-y$predy_cf_det)
  y$acc_cf <- 1-y$error_abs_cf/y$ZAC_HVTA
  
  bd <- union_all(bd, y)
  
}

rm(df); rm(n); rm(x.nn); rm(x.nn.pred); rm (y); rm(y.nn); rm(y.nn.pred)


# 4) Integración Pronóstico Agregado y Detallado----

#Calcular totales de la base de pronóstico detallado
fcst_ag <- aggregate(bd[,23], by = list(SUBREGION_CDA = bd$SUBREGION_CDA, agrupacion = bd$agrupacion,
                                          segmento_agrupado = bd$segmento_agrupado, marca = bd$marca,
                                          tamaño = bd$tamaño, retornable = bd$retornable, semana = bd$CALWEEK)
                      , FUN = sum, na.rm = TRUE) 

colnames(fcst_ag)[8] <- "predy_total"
fcst_ag$id <- paste(fcst_ag$SUBREGION_CDA, fcst_ag$agrupacion, fcst_ag$segmento_agrupado, fcst_ag$marca, fcst_ag$tamaño, fcst_ag$retornable, fcst_ag$semana)

bd$combo_ag_sem <- paste(bd$combo_ag, bd$CALWEEK)

#Calcular los pesos de distribución del nivel detallado
bd <- merge(bd, fcst_ag[,c(8,9)], by.x = "combo_ag_sem", by.y = "id")

bd$peso <- bd$predy_cf_det/bd$predy_total

bd <- merge(bd, x[,c(11,14)], by = "combo_ag_sem")
bd$predy_cf_ajust <- bd$predy_cf_ag * bd$peso
bd$error_abs_cf_ajust <- abs(bd$ZAC_HVTA-bd$predy_cf_ajust)
bd$acc_cf_ajust <- 1-bd$error_abs_cf_ajust/bd$ZAC_HVTA

###################################################

write.csv(x, file ="C:/Users/gcall/Downloads/Test Fcst/resultado_modelo_agregado.csv", row.names = FALSE)
write.csv(bd, file ="C:/Users/gcall/Downloads/Test Fcst/resultado_modelo_detallado.csv", row.names = FALSE)
