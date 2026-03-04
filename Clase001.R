# 1 Carga de librerias

library(tidyverse)   # dplyr + ggplot2 + tidyr + etc.
library(readxl)      # leer Excel
library(janitor)     # clean_names()
library(writexl)     # exportar a Excel
library(lubridate)   # trabajar con fechas: year(), ymd(), etc.

# 2 importación del Excel
bd_cruda <- read_excel(file.choose(), sheet = "Hoja1")

# 3 exploración inicial, primeras filas
head(bd_cruda)

# 4 exploración inicial, estructura con glimpse
glimpse(bd_cruda)

# 5 limpieza principal, construcción de bd_limpia
bd_limpia <- bd_cruda %>%
  clean_names() %>%
  mutate(
    across(where(is.character), ~ str_squish(.x))
  ) %>%
  select(-id, -estado, -pais_nace) %>%
  mutate(
    fecha = as.Date(fecha),
    anio  = year(fecha)
  )

# 6 verificar nombres nuevos
names(bd_limpia)

# 7 ver estructura final de bd_limpia
glimpse(bd_limpia)

# 8 exportar la base limpia
write_xlsx(bd_limpia, "BD_springfield_procesada.xlsx")

# 9 barrios con más registros
bd_limpia %>%
  count(barrio, sort = TRUE)

# 10 armas con porcentaje
bd_limpia %>%
  count(arma_empleada, sort = TRUE) %>%
  mutate(
    porcentaje = round(n / sum(n) * 100, 1)
  )

# 11 primer gráfico, top 10 barrios
bd_limpia %>%
  count(barrio, sort = TRUE) %>%
  slice_head(n = 10) %>%
  ggplot(aes(x = reorder(barrio, n), y = n)) +
  geom_col() +
  coord_flip() +
  labs(x = "Barrio", y = "Cantidad de eventos") +
  theme_minimal()

# 12 zona por clase de sitio
bd_limpia %>%
  count(zona, clase_sitio) %>%
  pivot_wider(names_from = zona, values_from = n, values_fill = 0)

# 13 jornada por arma empleada
bd_limpia %>%
  count(jornada, arma_empleada) %>%
  pivot_wider(names_from = arma_empleada, values_from = n, values_fill = 0)

# 14 sexo por clase de sitio
bd_limpia %>%
  count(sexo, clase_sitio) %>%
  pivot_wider(names_from = clase_sitio, values_from = n, values_fill = 0)

# 15 día por jornada
bd_limpia %>%
  count(dia, jornada) %>%
  pivot_wider(names_from = jornada, values_from = n, values_fill = 0)

# 16 rango etario por movilidad de la víctima MALO
bd_limpia <- bd_limpia %>%
  mutate(
    grupo_etario = case_when(
      !is.na(edad) & edad >= 0  & edad <= 18  ~ "0-18 Menores de edad",
      !is.na(edad) & edad >= 19 & edad <= 30  ~ "19-30 Jóvenes",
      !is.na(edad) & edad >= 31 & edad <= 59  ~ "31-59 Adultos",
      !is.na(edad) & edad >= 60 & edad <= 100 ~ "60-100 Adultos mayores",
      TRUE ~ NA_character_
    )
  )

bd_limpia %>%
  count(grupo_etario, movil_victima) %>%
  pivot_wider(names_from = movil_victima, values_from = n, values_fill = 0)

# 17: zona por movilidad del agresor
bd_limpia %>%
  count(zona, movil_agresor) %>%
  pivot_wider(names_from = movil_agresor, values_from = n, values_fill = 0)

# 18 profesión por arma empleada, sin romper la clase
tabla_prof_arma <- bd_limpia %>%
  count(profesion, arma_empleada) %>%
  pivot_wider(names_from = arma_empleada, values_from = n, values_fill = 0)

tabla_prof_arma %>%
  mutate(total = rowSums(across(where(is.numeric)))) %>%
  arrange(desc(total))

# =========================
# EXPORTAR TABLAS + GRÁFICOS A UN EXCEL (CON HOJAS NOMBRADAS)
# Requiere: bd_limpia ya creada en memoria
# =========================

# Paquetes
if (!requireNamespace("openxlsx", quietly = TRUE)) install.packages("openxlsx")
library(openxlsx)
library(tidyverse)

# 1) (Opcional pero recomendado) Crear grupo_etario si aún no existe
if (!("grupo_etario" %in% names(bd_limpia))) {
  bd_limpia <- bd_limpia %>%
    mutate(
      grupo_etario = case_when(
        !is.na(edad) & edad >= 0  & edad <= 18  ~ "0-18 Menores de edad",
        !is.na(edad) & edad >= 19 & edad <= 30  ~ "19-30 Jóvenes",
        !is.na(edad) & edad >= 31 & edad <= 59  ~ "31-59 Adultos",
        !is.na(edad) & edad >= 60 & edad <= 100 ~ "60-100 Adultos mayores",
        TRUE ~ NA_character_
      )
    )
}

# 2) Construcción de TABLAS (todas como data.frame para Excel)
tab_barrio_freq <- bd_limpia %>% count(barrio, sort = TRUE) %>% as.data.frame()

tab_arma_pct <- bd_limpia %>%
  count(arma_empleada, sort = TRUE) %>%
  mutate(porcentaje = round(n / sum(n) * 100, 1)) %>%
  as.data.frame()

tab_zona_clase <- bd_limpia %>%
  count(zona, clase_sitio) %>%
  pivot_wider(names_from = zona, values_from = n, values_fill = 0) %>%
  as.data.frame()

tab_jornada_arma <- bd_limpia %>%
  count(jornada, arma_empleada) %>%
  pivot_wider(names_from = arma_empleada, values_from = n, values_fill = 0) %>%
  as.data.frame()

tab_sexo_clase <- bd_limpia %>%
  count(sexo, clase_sitio) %>%
  pivot_wider(names_from = clase_sitio, values_from = n, values_fill = 0) %>%
  as.data.frame()

tab_dia_jornada <- bd_limpia %>%
  count(dia, jornada) %>%
  pivot_wider(names_from = jornada, values_from = n, values_fill = 0) %>%
  as.data.frame()

tab_etario_movil_victima <- bd_limpia %>%
  count(grupo_etario, movil_victima) %>%
  pivot_wider(names_from = movil_victima, values_from = n, values_fill = 0) %>%
  as.data.frame()

tab_zona_movil_agresor <- bd_limpia %>%
  count(zona, movil_agresor) %>%
  pivot_wider(names_from = movil_agresor, values_from = n, values_fill = 0) %>%
  as.data.frame()

tab_prof_arma_total <- bd_limpia %>%
  count(profesion, arma_empleada) %>%
  pivot_wider(names_from = arma_empleada, values_from = n, values_fill = 0) %>%
  mutate(total = rowSums(across(where(is.numeric)))) %>%
  arrange(desc(total)) %>%
  as.data.frame()

# 3) Gráfico(s): guardar PNG para insertarlo en Excel
plot_top10_barrio <- bd_limpia %>%
  count(barrio, sort = TRUE) %>%
  slice_head(n = 10) %>%
  ggplot(aes(x = reorder(barrio, n), y = n)) +
  geom_col() +
  coord_flip() +
  labs(x = "Barrio", y = "Cantidad de eventos") +
  theme_minimal()

plot_file_1 <- file.path(tempdir(), "plot_top10_barrios.png")
ggsave(filename = plot_file_1, plot = plot_top10_barrio, width = 10, height = 6, dpi = 200)

# 4) Crear libro Excel, hojas, tablas y gráficos
wb <- createWorkbook()

# Helper: nombres de hoja válidos (máx 31 caracteres, sin caracteres prohibidos)
sanitize_sheetname <- function(x) {
  x <- gsub("[:\\\\/\\?\\*\\[\\]]", " ", x)
  x <- str_squish(x)
  substr(x, 1, 31)
}

# Lista de tablas (nombre_hoja = objeto)
tablas <- list(
  "01_Frec_Barrio"            = tab_barrio_freq,
  "02_Arma_Porc"              = tab_arma_pct,
  "03_Zona_x_ClaseSitio"      = tab_zona_clase,
  "04_Jornada_x_Arma"         = tab_jornada_arma,
  "05_Sexo_x_ClaseSitio"      = tab_sexo_clase,
  "06_Dia_x_Jornada"          = tab_dia_jornada,
  "07_Etario_x_MovilVictima"  = tab_etario_movil_victima,
  "08_Zona_x_MovilAgresor"    = tab_zona_movil_agresor,
  "09_Prof_x_Arma_Total"      = tab_prof_arma_total
)

# Estilo básico de cabecera
headerStyle <- createStyle(textDecoration = "bold", halign = "center", valign = "center", border = "Bottom")

# Escribir cada tabla en su hoja
for (nm in names(tablas)) {
  sh <- sanitize_sheetname(nm)
  addWorksheet(wb, sh)
  writeData(wb, sh, tablas[[nm]], startRow = 1, startCol = 1, headerStyle = headerStyle)
  freezePane(wb, sh, firstRow = TRUE)
  setColWidths(wb, sh, cols = 1:ncol(tablas[[nm]]), widths = "auto")
}

# Hoja de gráficos
addWorksheet(wb, "10_Graficos")
insertImage(wb, "10_Graficos", file = plot_file_1, startRow = 2, startCol = 1, width = 10, height = 6, units = "in")
writeData(wb, "10_Graficos", data.frame(Grafico = "Top 10 Barrios (conteo)"), startRow = 1, startCol = 1, headerStyle = headerStyle)

# 5) Guardar archivo final
saveWorkbook(wb, "Resultados_Sesion1.xlsx", overwrite = TRUE)

#--------------------------------------------------------------------
#SLIDE 6 — Construir tabla para prueba

#Ejemplo zona - sexo

tabla_zona_sexo <- table(bd_limpia$zona, bd_limpia$sexo)

tabla_zona_sexo

#Ejemplo arma - sexo

tabla_arma_sexo <- table(bd_limpia$arma_empleada, bd_limpia$sexo)

tabla_arma_sexo

#SLIDE 7 — Ejecutar Chi-cuadrado
#Ejemplo zona - sexo
chisq.test(tabla_zona_sexo)

#Ejemplo arma - sexo
chisq.test(tabla_arma_sexo)

#SLIDE 8 Otro ejemplo Jornada x Arma
tabla_jornada_arma <- table(bd_limpia$jornada, bd_limpia$arma_empleada)

tabla_jornada_arma

chisq.test(tabla_jornada_arma)

#---------------------------------
#---------------------------------
# slide crear la variable binaria violento a partir de arma_empleada
bd_modelo <- bd_limpia %>%
  mutate(
    arma_empleada = str_squish(arma_empleada),
    violento = case_when(
      str_detect(str_to_lower(arma_empleada), "fuego") ~ 1,
      str_detect(str_to_lower(arma_empleada), "blanca") ~ 1,
      TRUE ~ 0
    )
  ) %>%
  filter(!is.na(jornada), !is.na(violento))

#SLIDE 3 — Código: revisar rápidamente que la variable quedó bien construida
bd_modelo %>% count(violento)
bd_modelo %>% count(jornada, violento)

#SLIDE 4 — Código: ajustar el modelo logístico con solo dos variables
modelo_logit <- glm(
  violento ~ jornada,
  data = bd_modelo,
  family = binomial(link = "logit")
)

summary(modelo_logit)

#SLIDE 5 — Código: convertir coeficientes a Odds Ratios e interpretarlos
odds_ratios <- exp(coef(modelo_logit))
odds_ratios

#SLIDE 6 — Código: poner Odds Ratios con intervalos de confianza para hablar con rigor
ic_or <- exp(confint(modelo_logit))
ic_or

#SLIDE 7 — Código: presentar resultados en una tabla limpia lista para informe
tabla_or <- tibble(
  termino = names(coef(modelo_logit)),
  beta = coef(modelo_logit),
  odds_ratio = exp(beta)
) %>%
  mutate(
    odds_ratio = round(odds_ratio, 3)
  )

tabla_or

# =========================================================
# EXPORTAR RESULTADOS A WORD (solo exportación)
# =========================================================

# Instalar si es necesario
if (!requireNamespace("officer", quietly = TRUE)) install.packages("officer")
if (!requireNamespace("flextable", quietly = TRUE)) install.packages("flextable")

library(officer)
library(flextable)

# Ruta Descargas
ruta_descargas <- file.path(Sys.getenv("HOME"), "Downloads")
if (!dir.exists(ruta_descargas)) {
  ruta_descargas <- file.path(Sys.getenv("USERPROFILE"), "Downloads")
}

archivo_word <- file.path(ruta_descargas, "Sesion2_Resultados_Logit.docx")

# Crear documento
doc <- read_docx()

# ----------------------------
# Agregar contenido al Word
# ----------------------------

doc <- doc %>%
  body_add_par("SESIÓN 2 — RESULTADOS REGRESIÓN LOGÍSTICA", style = "heading 1")

# Slide 3 – Tablas descriptivas
doc <- doc %>%
  body_add_par("Distribución de la variable dependiente (Violento)", style = "heading 2") %>%
  body_add_flextable(flextable(tab_violento) %>% autofit()) %>%
  body_add_par("Distribución por Jornada", style = "heading 2") %>%
  body_add_flextable(flextable(tab_jornada_violento) %>% autofit())

# Slide 4 – Resumen del modelo
doc <- doc %>%
  body_add_par("Resumen del modelo logístico", style = "heading 2") %>%
  body_add_par(paste(capture.output(summary(modelo_logit)), collapse = "\n"), style = "Normal")

# Slide 5 – Odds Ratios
if (!exists("tab_or_simple")) {
  tab_or_simple <- tibble(
    termino = names(exp(coef(modelo_logit))),
    odds_ratio = as.numeric(exp(coef(modelo_logit)))
  ) %>%
    mutate(odds_ratio = round(odds_ratio, 3)) %>%
    as.data.frame()
}

# Slide 6 – Intervalos de confianza
if (!exists("tab_ic_or")) {
  tab_ic_or <- as.data.frame(exp(confint(modelo_logit)))
  colnames(tab_ic_or) <- c("LI_95_OR", "LS_95_OR")
  tab_ic_or <- tibble(
    termino = rownames(tab_ic_or),
    LI_95_OR = tab_ic_or$LI_95_OR,
    LS_95_OR = tab_ic_or$LS_95_OR
  ) %>%
    mutate(
      LI_95_OR = round(LI_95_OR, 3),
      LS_95_OR = round(LS_95_OR, 3)
    ) %>%
    as.data.frame()
}

# Slide 7 – Tabla ejecutiva
doc <- doc %>%
  body_add_par("Tabla Ejecutiva Final", style = "heading 2") %>%
  body_add_flextable(flextable(tabla_or) %>% autofit())

# Guardar documento
print(doc, target = archivo_word)

cat("Documento guardado en:", archivo_word)



