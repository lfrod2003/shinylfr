# =============================================================================
#                     Indicadores de Calidad de Predicción
# =============================================================================

# Carga de librerías ------------------------------------------------------
options(useFancyQuotes = FALSE)
library(RODBC)          # Conexión a la base de datos
library(RODBCext)       # Conexión a la base de datos
library(config)
library(shiny)
library(rhandsontable)  # Visualización tablas
library(htmlwidgets)
library(scales)
library(lubridate)
library(shinydashboard)
library(DT)
library(ggplot2)
library(plotly)         # Graficas dinámicas plotly
library(RCurl)
library(shinyjs)
library(shinycssloaders)# Reloj para tiempo de espera
library(shinyWidgets)
#library(googleway)      # Presentación de ubicación del punto de consumo en el mapa. Ref absoluta a add_markers
library(plyr)
library(dplyr)          # Adecuacion de tablas para visualización
library(googleway)      # Presentación de ubicación del punto de consumo en el mapa. Ref absoluta a add_markers
library(stringr)
require(colorspace)
library(shinyBS)        # Utilizada para mostrar tooltip
#library(rintrojs)       # Presentación de ayuda en pantalla
library(openxlsx)




source("helpers.R")
source("../_shared/_filtros.R")
#reactlog_enable()

Sys.setenv(LANGUAGE="ES")
options(encoding = 'UTF-8')


# función para convertir columna de datos jpg a cod64 
aCode64 <- function(objetoJPG) {
  base64Encode(objetoJPG,"text")
}

# Retornar tabla para despliegue ---------------------------------------------
tabla.despliegue <- function(tabla.temp, obj) {
  
  if(is.null(tabla.temp)){
    DF <- data.frame(character = c())
    rhandsontable(DF,
                  colHeaders = c(
                    ''
                  ),
                  readOnly = TRUE
    )
  }else{
    rhandsontable(tabla.temp[,2:8],
                  rowHeaders = TRUE,
                  height =210,
                  search = FALSE,
                  readOnly = TRUE
                  ,selectCallback = TRUE)%>% 
      #hot_col(1, halign = "htCenter", readOnly = FALSE) %>% 
      hot_col(7,format="#.0%", halign="htRight" ) %>% 
      hot_heatmap(7, color_scale =c("#ffffff", "#ffffff")) %>%  # Habilitar escala de color
      hot_cols(columnSorting = TRUE)  %>%   
      hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE) %>%    # Bloquea opciones de menú contextual
      hot_table(highlightCol = TRUE, highlightRow = TRUE                 # Resalta fila y columna seleccionada
      )
  }
  
 
}

tablaLector.despliegue <- function(tabla.temp) {
  tablaexportar <<- tabla.temp
  
  rhandsontable(tabla.temp,
                rowHeaders = TRUE,
                height =380,
                search = FALSE,
                readOnly = TRUE
                ,selectCallback = TRUE)%>% 
    #hot_col(5,format="dd-MM-yyyy" ) %>% 
    hot_cols(columnSorting = TRUE)  %>%   
    hot_context_menu(allowRowEdit = FALSE, allowColEdit = FALSE) %>%    # Bloquea opciones de menú contextual
    hot_table(highlightCol = TRUE, highlightRow = TRUE                 # Resalta fila y columna seleccionada
    )
}

# Retornar mapa para despliegue --------------------------------------------
mapa.despliegue<- function(tabla.temp, obj, datos.consulta, banderaRefresco,tabla.datos.valida, input) { #} datos.imagen, input) { 

  if (tabla.datos.valida) {
    tabla.temp <-  hot_to_r(input$datos)
    ordenes <- subset(tabla.temp, tabla.temp[,1] == TRUE)[,-1]
  } else if(nrow(obj)!=0) {
    ordenes <- obj
  } else {
    ordenes <- subset(tabla.temp, tabla.temp[,1] == TRUE)[,-1]
  }


  for(i in 1:nrow(ordenes)){                                      # Extrae el número de cliente del hipervínculo
    pos.cliente <- gregexpr(">",ordenes$CodigoPuntoConsumo[i])
    pos.cliente <- pos.cliente[[1]][1] + 1
    fin.cad <- nchar(ordenes$CodigoPuntoConsumo[i]) - 4
    ordenes$CodigoPuntoConsumo[i] <- substr(ordenes$CodigoPuntoConsumo[i],pos.cliente,fin.cad)
  }
  ordenes$CodigoPuntoConsumo <- unlist(ordenes$CodigoPuntoConsumo)
  
  datos<-left_join(ordenes,datos.consulta,by = "CodigoPuntoConsumo") # Efectua unión para traer coordenadas
  datos.consulta$PromedioActiva6CN<-as.numeric(as.character(datos.consulta$PromedioActiva6CN))  # Convierte a numérico
  datos.consulta$PromedioActiva6CN[is.na(datos.consulta$PromedioActiva6CN)]=0                  # Ajuste de valores nulos
  datos$LatitudPuntoConsumo <-as.numeric(as.character(datos$LatitudPuntoConsumo))
  datos$LongitudPuntoConsumo <-as.numeric(as.character(datos$LongitudPuntoConsumo))

  
  if (banderaRefresco) {
    google_map_update(map_id = "map") %>%
      clear_markers() %>%
      googleway::add_markers(data = datos,lat = "LatitudPuntoConsumo", lon = "LongitudPuntoConsumo", 
                             info_window = "Texto1",close_info_window = TRUE)
                             #info_window = "texto",close_info_window = TRUE)
  } else {
    google_map(key = api_key, data = datos, style=estilo.map01) %>%
      googleway::add_markers(lat = "LatitudPuntoConsumo", lon = "LongitudPuntoConsumo", 
                             info_window = "Texto1",close_info_window = TRUE)
  #  info_window = "texto",close_info_window = TRUE)
  }
}


# Configuración de bd y código para api google --------------------------------------------------------------
configuracion.conexion <<- 'externalWM' # Sys.getenv("CONEXIONSHINY") #'windows', 'externalENERG' para energuate
#  'GoogleKey' CodigoParametro api google en [DWSRBI_KRONOS].[Aux].[Parametros]
config <- config::get(config=configuracion.conexion)
config.conexion <- config$conexion
conexion <- odbcDriverConnect (config.conexion)
cad.sql <- "SELECT TOP 1 ValorParametro FROM [DWSRBI_KRONOS].[Aux].[Parametros] where EstadoParametro = 1 AND   CodigoParametro = 'GoogleKey'"
api_key <-sqlExecute(channel = conexion, query = cad.sql,fetch= T,as.is=T)
odbcClose(conexion)
api_key <<- api_key[1,1]


# Contadores y llamadas --------------------------------------------------------------
ord.gen <<- 0  # Contador para presentar la cantidad de órdenes generadas en la sesión
analista <<- "NBC"                     # Cadena para identificación del solicitante
#api_key <- "AIzaSyAYoARt3zU_arrmeiZmMCjXhLEyRoE7Z2Q"  # Solicitar clave a WM y cambiar
Sys.setenv(LANGUAGE="ES")
options(encoding = 'UTF-8')
locale <- Sys.getlocale(category="LC_COLLATE")
if (grepl("Spanish", locale, fixed=TRUE)) {
  separador.csv <<- ';'
} else {
  separador.csv <<- ','
}

# Captura de parámetro para URL servidor shiny -------------
config <- config::get(config=configuracion.conexion)
config.conexion <- config$conexion
conexion <- odbcDriverConnect (config.conexion)
cad.sql<- "SELECT [ValorParametro]  FROM [DWSRBI_KRONOS].[Aux].[Parametros]  WHERE [CodigoParametro] = 'URLshiny'"
URL.servidor.shiny <-sqlExecute(channel = conexion, query = cad.sql,fetch= T,as.is=T)
URL.servidor.shiny <- URL.servidor.shiny [1,1]
str.http <<- paste(URL.servidor.shiny,"hvurl/?Codigo=",sep="") # Cadena de llamada para vinculo a hoja de vida




# Captura de campañas disponibles en BD -----------------------------------
cad.sql<- "[dbo].[Leer_Campanas1]"
nom.camp <- sqlExecute(channel = conexion, query = cad.sql,fetch= T,as.is=T)
colnames(nom.camp)<-"Nombre"
nom.camp <- nom.camp  %>% filter(!is.na(Nombre))

# cad.sql <- "SELECT DISTINCT [NombreServicio] FROM [Dimension].[Servicio] where codigoservicio > 0  ORDER BY [NombreServicio]"
# nom_servicios <- sqlExecute(channel = conexion, query = cad.sql,fetch= T,as.is=T)
# colnames(nom_servicios) <- "Nombre"
# nom_servicios <- nom_servicios   %>% filter(!is.na(Nombre))
# nom_servicios <- c("G. clientes", "PIMT", "Masivos", "Consumos fijos", "Autoproductores")

# Consultar datos de filtros
consultarDatosFiltros(conexion)

# Carga contenido filtros
cargarFiltrosIniciales()

odbcClose(conexion)

legalizacionLista = format(seq(as.Date("2022-01-01"), today(), by = "months"),"%Y%m")
#aniosLegalizacion_lista <- "2019":format(today(), "%Y")
#mesesLegalizacion_lista <- str_pad("01":"12", 2, pad = "0")

# font negro para pickerinput
col_list2 <- c("black")
colorspi <- paste0("color:",rep(c('black'),each=10),";")

ui = dashboardPage(
  dashboardHeader(title = "Indicadores de Calidad de Predicción",
                  tags$li(div(
                    img(src = 'Kronos.png',
                        title = "Indicadores de Calidad de Predicción", height = "30px"),
                    style = "padding-top:10px; padding-bottom:10px; margin-right:10px;"),
                    class = "dropdown"),
                  dropdownMenuOutput("MensajeOrdenes")),
  
  # Sidebar -----------------------------------------------------------------
  
  dashboardSidebar(width = 500,
                   
                   # Codigo para reducir espacio entre objetos Shiny                   
                   tags$head(
                     tags$style(
                       HTML(".form-group {
                            margin-bottom: 0 !important;
                            }"))),
                   
                   # Codigo para no mostrar errores en interfaz Shiny
                   tags$style(type="text/css",
                              ".shiny-output-error { visibility: hidden; }",
                              ".shiny-output-error:before { visibility: hidden; }"
                   ),
                   fluidRow( column(width = 6,offset = 0, style='padding:0px;',
                                    box(id = "Comandos", width = 12, 
                                        status = NULL,  
                                        background = "black",
                                        fluidRow( 
                                              column(width = 2,offset = 1,
                                                         actionButton("ReiniciarControles", label = icon("fas fa-sync"),
                                                                      style="color: #fff; background-color: #0070ba; border-color: #0070ba"),
                                                         bsTooltip("ReiniciarControles", "Reiniciar valores de filtro", placement = "bottom", trigger = "hover", options = NULL)
                                              ),
                                              column(width = 2,
                                                     offset = 1,
                                                     actionButton("TraerDatos", label = icon("fas fa-play"),
                                                                  style="color: #fff; background-color: #0070ba; border-color: #0070ba"), 
                                                     bsTooltip("TraerDatos", "Ejecutar consulta", placement = "bottom", trigger = "hover", options = NULL)
                                              )
                                        )
                                        
                                    ),
                                    
                                    box(id = "FiltroLegalizacion", width = 12, status = NULL,  background = "black",
                                        
                                        pickerInput("periodo_legal","Periodo de legalización:",
                                                    choices = legalizacionLista,
                                                    selected = NULL,
                                                    multiple=T,
                                                    choicesOpt = list(
                                                      style=rep(paste0("color:black;"),length(legalizacionLista))),
                                                    options = list(
                                                      `none-selected-text` = "Periodo legal"
                                                    )),
                                        
                                        radioButtons(inputId = "pendientes_legal", 
                                                     label = "Pendientes de legalización:", 
                                                     choices = c("Si", "No"),                                  
                                                     inline = TRUE
                                        )
                                    ),
                                    
                                    obtenerFiltrosZona(),
                                    obtenerFiltrosDepartamento()
                              ),
                             column(width = 6, offset = 0, style='padding:0px;',
                                    obtenerFiltrosEmpresa()
                             ))
  ),
  
  # Body --------------------------------------------------------------------
  ##00828F;
  dashboardBody(    
    tags$head(tags$style(HTML('
                              
                              /* Separacion entre objetos */
                              .form-group {
                              margin-bottom: 0 !important;
                              }
                              
                              /* logo */
                              .skin-blue .main-header .logo {
                              background-color: #0070ba;
                              }
                              
                              /* logo when hove red */
                              .skin-blue .main-header .logo:hover {
                              background-color: #0F3D3F;
                              }
                              
                              # /* main sidebar */
                              # .skin-blue .main-sidebar {
                              # background-color: #0070ba;
                              # }
                              
                              /* navbar (rest of the header) */
                              .skin-blue .main-header .navbar {
                              background-color: #0070ba;
                              }
                              
                              /* body */
                              .content-wrapper, .right-side {
                              background-color: #FFFFFF;
                              }
                              
                              /* color para botones modales */
                                #modal1 button.btn.btn-default {
                                color: #fff; background-color: #0070ba; border-color: #0070ba
                                }'
                              
    )
    )
    ), # Fin estilo

    
    # Paneles -----------------------------------------------------------------
    
    tabsetPanel( type = "tabs", 
                              
                 tabPanel("Datos",
                          icon = icon("fas fa-table"),
                          hr(),
                          useShinyjs(),
                          fluidRow( 
                            box(#height = 420,
                              width = 12,
                              status = "warning",
                              solidHeader = FALSE,
                              title = "Indicadores de Calidad de Predicción",
                              br(),
                              radioButtons("orden_tabla","Ordenar por:", c("Bolsa"="Bolsa","Pendientes"="Pendientes","Cerradas"="Cerradas", "Normalizadas"= "Normalizadas","CNR"="CNR","Fallidas"="Fallidas"), inline=T),
                              br(),
                              withSpinner(DT::dataTableOutput("datos", height = "200px"),
                                          color = getOption("spinner.color", default = "#0070ba")),
                              # withSpinner(rHandsontableOutput("datos", height = "600px"),         # Incluir spinner de espera
                              #             color = getOption("spinner.color", default = "#0070ba"),
                              #             id = "contenedorTabla1"
                              # ),
                              br(),
                              div(rHandsontableOutput("datosLector"),
                                 #         color = getOption("spinner.color", default = "#0070ba"),
                                          id = "contenedorTabla2"
                              )
                              
                            )),
                          hr(),
                          fluidRow(column(width=12,
                                          plotlyOutput("grafico", height = "200px"))),
                          fluidRow(hr() ,      
                                   
                                   # column(width = 2,
                                   #        actionButton("MostrarTabla",   # Se crea el botón deshabilitado para que
                                   #                     "Actualizar tabla"    # sea habilitado posteriormente por código
                                   #        )),
                                   # column(width = 2,
                                   #        offset = 6,
                                   #        shinyjs::disabled(downloadButton('Descargar', 'Descargar excel',
                                   #                                         style="color: #fff; background-color: #0070ba; border-color: #0070ba"),  # Se crea el botón deshabilitado para que
                                   #                          bsTooltip("Descargar", "Descargar usuarios seleccionados a archivo excel", 
                                   #                                    placement = "bottom", trigger = "hover", options = NULL)
                                   #        )                                           # sea habilitado posteriormente por código
                                   #        
                                   # ),
                                   # column(width = 2,
                                   #        shinyjs::disabled(  # Se crea el botón deshabilitado para que
                                   #          actionButton("Guardar",   
                                   #                       "Generar órdenes",   
                                   #                       icon = icon("fas fa-save"),
                                   #                       style="color: #fff; background-color: #0070ba; border-color: #0070ba"
                                   #          ),
                                   #          bsTooltip("Guardar", "Crear órdenes de inspección para usuarios seleccionados", 
                                   #                    placement = "bottom", trigger = "hover", options = NULL)
                                   #        )
                                   # )
                          ),
                          textOutput("LlavePC"),
                          tags$head(tags$style("#LlavePC{color: white;
                                 font-size: 20px;
                                 font-style: italic;
                                 }"
                          )
                          )                         
                 ),
                 tabPanel("Ayuda",
                          icon = icon("fas fa-table"),
                          includeMarkdown("Ayuda.md")
                          
                 ),
                 
                 id="TabsApp"
    )
  )
)

# Server ------------------------------------------------------------------

server <- shinyServer(function(input, output, session) { # Importante iniciar con shinyServer para que funcione la ayuda
  
  updatePickerInput(session,"estado",selected = '')
  updatePickerInput(session,"tarifa",selected = '')
  updatePickerInput(session,"mercado",selected = '')
  
  click("TraerDatos")
  
  consulta.activa <<- FALSE
  datos.consulta <- NA
  tipo.grafica <- 'P'
  tabla.datos.valida <- FALSE
  valores_validos_zona <<- NA
  valores_validos_region <<- NA
  valores_validos_centro <<- NA
  valores_validos_Geografia_departamento  <<- NA
  valores_validos_Geografia_municipio <<- NA

  valores_validos_estado <<- NA
  valores_validos_tarifa <<- NA
  valores_validos_mercado <<- NA
  
  # Ocultar Botones auxiliares para sincronizar paneles, paneles
  shinyjs::hide("MostrarTabla")
  shinyjs::hide("MostrarMapa")
  
  tabla.temp <<- hot_to_r(NULL)    # Tabla temporal para alimentar rhandsontable
  
  # output$datos <- renderRHandsontable({   # Actualiza rhandsontable
  #   tabla.despliegue(tabla.temp, data.grafica()$sel)
  # })
  
  # Ayuda -------------------------------------------------------------------
  
  
  # Cambio de tab ----------------------------------------------------------------------
  observeEvent(input$TabsApp, {
    if (input$TabsApp == "Datos") {
      tabla.datos.valida <<- TRUE
    } else if (input$TabsApp == "Mapa") {
      click("MostrarMapa")
    }
  }, ignoreInit = TRUE)
  
  
  # Seleccion de fila en tabla de datos --------------------------------------------------
  observeEvent(
    input$datos_sort
    ,{

      xyz <<- input$datos_sort$data
    }
  )
  
  
  # Carga de datos a listas desplegables
  manejarEventosFiltros(input, output, session)
  
  # Reinicial controles
  manejarReinicioFiltros(input, output, session)
  
  observeEvent(input$ReiniciarControles, {
    updatePickerInput(session,"estado",selected = '')
    updatePickerInput(session,"tarifa",selected = '')
    updatePickerInput(session,"mercado",selected = '')
  }, ignoreInit = TRUE)
  
  tabla_despliegue <- function() {
    datos.consulta <- datos.consulta[
      order( datos.consulta[,1], datos.consulta[,2] , -datos.consulta[,match(input$orden_tabla,names(datos.consulta))] ),
    ]
    
    row.names(datos.consulta) <- NULL
    
    tabla.datos <- datos.consulta
    
    tabla.datos$`NombreCampaña` <- ifelse(tabla.datos$AuxOrder=='zzzz',paste0("<b>",tabla.datos$`NombreCampaña`, "</b>"), tabla.datos$`NombreCampaña`)
    tabla.datos$`NombreTipoCampaña` <- ifelse(tabla.datos$AuxOrder=='zzzz',paste0("<b>",tabla.datos$`NombreTipoCampaña`, "</b>"), tabla.datos$`NombreTipoCampaña`)
    tabla.datos$Bolsa <- ifelse(tabla.datos$AuxOrder=='zzzz',paste0("<b>",tabla.datos$Bolsa, "</b>"), tabla.datos$Bolsa)
    tabla.datos$Pendientes <- ifelse(tabla.datos$AuxOrder=='zzzz',paste0("<b>",tabla.datos$Pendientes, "</b>"), tabla.datos$Pendientes)
    tabla.datos$Normalizadas <- ifelse(tabla.datos$AuxOrder=='zzzz',paste0("<b>",tabla.datos$Normalizadas, "</b>"), tabla.datos$Normalizadas)
    tabla.datos$Cerradas <- ifelse(tabla.datos$AuxOrder=='zzzz',paste0("<b>",tabla.datos$Cerradas, "</b>"), tabla.datos$Cerradas)
    tabla.datos$Fallidas <- ifelse(tabla.datos$AuxOrder=='zzzz',paste0("<b>",tabla.datos$Fallidas, "</b>"), tabla.datos$Fallidas)
    tabla.datos$CNR <- ifelse(tabla.datos$AuxOrder=='zzzz',paste0("<b>",tabla.datos$CNR, "</b>"), tabla.datos$CNR)
    tabla.datos$PrcAvanceText <- ifelse(tabla.datos$AuxOrder=='zzzz',paste0("<b>",tabla.datos$PrcAvanceText, "</b>"), tabla.datos$PrcAvanceText)
    tabla.datos$PrcNormalizacionText <- ifelse(tabla.datos$AuxOrder=='zzzz',paste0("<b>",tabla.datos$PrcNormalizacionText, "</b>"), tabla.datos$PrcNormalizacionText)
    tabla.datos$PrcCNRText <- ifelse(tabla.datos$AuxOrder=='zzzz',paste0("<b>",tabla.datos$PrcCNRText, "</b>"), tabla.datos$PrcCNRText)
    tabla.datos$PrcFallidasText <- ifelse(tabla.datos$AuxOrder=='zzzz',paste0("<b>",tabla.datos$PrcFallidasText, "</b>"), tabla.datos$PrcFallidasText)
    
    tabla.datos <- subset(tabla.datos, 
                          select = c("Tipo_generacion", "NombreTipoCampaña",
                                     "Bolsa", "Pendientes","Cerradas","Normalizadas", "CNR","Fallidas", 
                                     "PrcAvanceText", "PrcNormalizacionText", "PrcCNRText", 
                                     "PrcFallidasText"))
    
    tabla.datos <<- tabla.datos
    
    output$datos <- DT::renderDataTable({
      DT::datatable(
        tabla.datos, 
        colnames=c("",  "Nombre grupo", "Bolsa", "Pendientes","Cerradas","Normalizadas", "CNR","Fallidas"
                   , "Avance(%)", "Norm(%)", "CNR(%)", "Fallida(%)"),
        class="compact",
        rownames = FALSE,
        extensions = 'RowGroup',
        options = list(
          rowGroup = list(dataSrc=c(0)),
          # columnDefs = list(list(visible=FALSE, targets=c(0))),
          columnDefs = list(list(className = 'dt-right', targets = 2:11)),
          scrollX = TRUE,
          scroller = TRUE,
          pageLength = 100,
          # order = list(0, 'desc', 1, 'asc')
          ordering = FALSE
        ), 
        escape = FALSE
      )
    })
  }
  
  
  observeEvent(input$orden_tabla, {
    tabla_despliegue()
  }, ignoreInit = TRUE)
  
  observeEvent(input$ReiniciarControles, {
    isolate(
      updatePickerInput(session,
                        "periodo_legal",
                        selected = ''))
    
  }, ignoreInit = TRUE)
  
  
  # Hacer consulta basada en valores de filtros  ---------------------------
  observeEvent(input$TraerDatos, {
    noHayValores <- FALSE
    fecha_minimo <- today() %m-% months(2)
    periodoSel <<- paste0(substring(fecha_minimo,1,4),substring(fecha_minimo,6,7))
    condiciones_consulta = ' 1>0 '
    
    pendientes_legal <- input$pendientes_legal
    periodo_legal <- input$periodo_legal
    if(!is.null(periodo_legal)){
      periodos <- gsub(',','',periodo_legal)
      condicion1 <- sprintf("i.LlaveFechaEjecucion > %s00 AND ",periodos)
      condicion2 <- sprintf("i.LlaveFechaEjecucion < %s40",periodos)
      condiciones_consulta <- paste0(condiciones_consulta, " AND ", paste(paste("(",condicion1, condicion2,")"), collapse = ' OR '))
      
      if(pendientes_legal == "Si"){
        condiciones_consulta <- paste0("(",condiciones_consulta, ") OR i.LlaveFechaEjecucion IS NULL")
      }else{
        condiciones_consulta <- paste0("(",condiciones_consulta, ") AND i.LlaveFechaEjecucion IS NOT NULL")
      }
      
      condiciones_consulta <- paste0("( ",condiciones_consulta, " )")
    }else{
      if(pendientes_legal != "Si"){
        condiciones_consulta <- paste0(condiciones_consulta, " AND i.LlaveFechaEjecucion IS NOT NULL")
      }
    }
    
    if(!is.null(input$zona_comercial)){
      # Construir condiciones de consulta:
      if (seleccion_operativa != "no") {
        if (nrow(valores_validos) == 0) {
          noHayValores <- TRUE
        } else { 
      
          #f1 <- substr(gsub("[\r\n]", "",paste0(unique(valores_validos[c("LlaveZona")]), collapse=",")),1,stop =1000000L)
          valores_validos <- consulta_zona(seleccion_operativa, valores_ZonaOperativa,input)
          
          f1 <- t(unique(valores_validos[c("LlaveZona")]))
          f1 <- paste0(apply(f1, 1, function(x) paste0(x)), collapse=",")
          #f1 <- paste0(gsub(":",",",unique(valores_validos[c("LlaveZona")])), collapse=",")
          f1 <- str_replace(f1, "[(]", "")
          f1 <- str_replace(f1, "[c]", "")
          f1 <- str_replace(f1, "[)]", "")
          condiciones_consulta <- gsub("[\r\n]", "",paste0(condiciones_consulta,
                                                           " AND l.Llavezona IN ( ",substr(f1,1,stop =1000000L)," ) ") )
        }
      }
    }
    
    if(!is.null(input$region)){
      if (seleccion_geografia != "no") {
        if (nrow(valores_validos_Geografia) == 0) {
          noHayValores <- TRUE
        } else {
          #f1 <- substr(gsub("[\r751:\n]", "",paste0(unique(valores_validos_Geografia[c("LlaveGeografia")]), collapse=",")),1,stop =1000000L)
          valores_validos_Geografia <- consulta_geografia(seleccion_geografia, valores_Geografia,input)
          f1 <- t(unique(valores_validos_Geografia[c("LlaveGeografia")]))
          f1 <- paste0(apply(f1, 1, function(x) paste0(x)), collapse=",")
          #f1 <- paste0(gsub(":",",",unique(valores_validos_Geografia[c("LlaveGeografia")])), collapse=",")
          f1 <- str_replace(f1, "[(]", "")
          f1 <- str_replace(f1, "[c]", "")
          f1 <- str_replace(f1, "[)]", "")
          condiciones_consulta <- gsub("[\r\n]", "",paste0(condiciones_consulta,
                                                           " AND l.LlaveGeografia IN ( ",substr(f1,1,stop =1000000L)," ) ") )
        }
      }
    }
    
    if(!is.null(input$circuitos)){
      if (seleccion_circuito != "no") {
        if (nrow(valores_validos_circuitos) == 0) {
          noHayValores <- TRUE
        } else {
          f1 <- t(unique(valores_validos_circuitos[c("LlaveCircuito")]))
          f1 <- paste0(apply(f1, 1, function(x) paste0(x)), collapse=",")
          #f1 <- paste0(gsub(":",",",unique(valores_validos_circuitos[c("LlaveCircuito")])), collapse=",")
          f1 <- str_replace(f1, "[(]", "")
          f1 <- str_replace(f1, "[c]", "")
          f1 <- str_replace(f1, "[)]", "")
          condiciones_consulta <- gsub("[\r\n]", "",paste0(condiciones_consulta,
                                                           " AND l.LlaveCircuito IN ( ",substr(f1,1,stop =1000000L)," ) ") )
        }
      }
    }
    
    if(!is.null(input$iniciativa)){
      if (seleccion_iniciativa != "no") {
        f1 <- t(unique(valores_validos_iniciativa[c("LlaveCampana")]))
        f1 <- paste0(apply(f1, 1, function(x) paste0(x)), collapse=",")
        f1 <- str_replace(f1, "[(]", "")
        f1 <- str_replace(f1, "[c]", "")
        f1 <- str_replace(f1, "[)]", "")
        condiciones_consulta <- gsub("[\r\n]", "",paste0(condiciones_consulta,
                                                         " AND c.LlaveCampaña IN ( ",substr(f1,1,stop =1000000L)," ) ") )
      }
    }
    
    c1 <- ''
    c2 <- ''
    
    if(!is.null(input$cnae)){
      if (seleccion_CNAE != 'no') {
        if (nrow(valores_validos_cnae_division) == 0) {
          noHayValores <- TRUE
        } else {
          f1 <- t(unique(valores_validos_cnae_division[c("LlaveActividadEconomica")]))
          f1 <- paste0(apply(f1, 1, function(x) paste0(x)), collapse=",")
          #f1 <- paste0(gsub(":",",",unique(valores_validos_cnae_division[c("LlaveActividadEconomica")])), collapse=",")
          f1 <- str_replace(f1, "[(]", "")
          f1 <- str_replace(f1, "[c]", "")
          c1 <- str_replace(f1, "[)]", "")
          
          # condiciones_consulta <- gsub("[\r\n]", "",paste0(condiciones_consulta,
          #                                                  " AND l.LlaveActividadEconomica IN ( ",substr(f1,1,stop =1000000L)," ) ") )
        }
      }
    }
    
    if(!is.null(input$cnaegrupo)){
      if (seleccion_grupo != 'no') {
        
        if (nrow(valores_validos_CNAEgrupo) == 0) {
          noHayValores <- TRUE
        } else {
          f1 <- t(unique(valores_validos_CNAEgrupo[c("LlaveActividadEconomica")]))
          f1 <- paste0(apply(f1, 1, function(x) paste0(x)), collapse=",")
          f1 <- paste0(gsub(":",",",unique(valores_validos_CNAEgrupo[c("LlaveActividadEconomica")])), collapse=",")
          f1 <- str_replace(f1, "[(]", "")
          f1 <- str_replace(f1, "[c]", "")
          c2 <- str_replace(f1, "[)]", "")
          # condiciones_consulta <- gsub("[\r\n]", "",paste0(condiciones_consulta,
          #                                                  " AND LlaveActividadEconomica IN ( ",substr(f1,1,stop =1000000L)," ) ") )
        }
      }
    }
    
    if(!is.null(input$cnaegrupo)){
      # Unificar seleccion_CNAE y seleccion_grupo
      if (seleccion_CNAE != 'no' && seleccion_grupo != 'no') {
        condiciones_CNAE <- gsub("[\r\n]", "",substr(c1,1,stop =1000000L),  
                                                         ",",substr(c2,1,stop =1000000L) )
        condiciones_consulta <- paste0(condiciones_consulta,
                                       " AND l.LlaveActividadEconomica IN ( ",condiciones_CNAE," ) ")
      } else if (seleccion_CNAE != 'no' ) {
        condiciones_CNAE <- gsub("[\r\n]", "",substr(c1,1,stop =1000000L) )
        condiciones_consulta <- paste0(condiciones_consulta,
                                       " AND l.LlaveActividadEconomica IN ( ",condiciones_CNAE," ) ")
                                                         
        
      } else if (seleccion_grupo != 'no') {
        condiciones_CNAE<- gsub("[\r\n]", "",substr(c2,1,stop =1000000L) )
        condiciones_consulta <- paste0(condiciones_consulta,
                                       " AND l.LlaveActividadEconomica IN ( ",condiciones_CNAE," ) ")
      }
    }

    if(!is.null(input$estado)){
      if (seleccion_estado != 'no') {
        ifelse (is.na(valores_validos_estado),
                {
                  lista_seleccion <- input$estado
                  valores_validos_estado <<- valores_Estados %>%
                    filter(valores_Estados$NombreEstado %in% lista_seleccion)
                },NA)  
        if (nrow(valores_validos_estado) == 0) {
          noHayValores <- TRUE
        } else {
          f1 <- t(unique(valores_validos_estado[c("LlaveEstado")]))
          f1 <- paste0(apply(f1, 1, function(x) paste0(x)), collapse=",")
          #f1 <- paste0(gsub(":",",",unique(valores_validos_tension[c("Llavetension")])), collapse=",")
          f1 <- str_replace(f1, "[(]", "")
          f1 <- str_replace(f1, "[c]", "")
          f1 <- str_replace(f1, "[)]", "")
          # condiciones_consulta <- gsub("[\r\n]", "",paste0(condiciones_consulta,
          #                                                  " AND LlaveEstrato IN ( ",substr(f1,1,stop =1000000L)," ) ") )
        }
      }
    }
    
    if(!is.null(input$tension)){
      if (seleccion_tension != 'no') {
        if (nrow(valores_validos_tension) == 0) {
          noHayValores <- TRUE
        } else {
          f1 <- t(unique(valores_validos_tension[c("Llavetension")]))
          f1 <- paste0(apply(f1, 1, function(x) paste0(x)), collapse=",")
          #f1 <- paste0(gsub(":",",",unique(valores_validos_tension[c("Llavetension")])), collapse=",")
          f1 <- str_replace(f1, "[(]", "")
          f1 <- str_replace(f1, "[c]", "")
          f1 <- str_replace(f1, "[)]", "")
          condiciones_consulta <- gsub("[\r\n]", "",paste0(condiciones_consulta,
                                                           " AND l.LlaveTension IN ( ",substr(f1,1,stop =1000000L)," ) ") )
        }
      }
    }
    
    if(!is.null(input$tarifa)){
      if (seleccion_tarifa != 'no') {
        ifelse (is.na(valores_validos_tarifa),
                {
                  lista_seleccion <- input$tarifa
                  valores_validos_tarifa <<- valores_Tarifa %>%
                    filter(valores_Tarifa$NombreTarifa %in% lista_seleccion)
                } , NA) 
        if (nrow(valores_validos_tarifa) == 0) {
          noHayValores <- TRUE
        } else {
          f1 <- t(unique(valores_validos_tarifa[c("LlaveTarifa")]))
          f1 <- paste0(apply(f1, 1, function(x) paste0(x)), collapse=",")
          #f1 <- paste0(gsub(":",",",unique(valores_validos_tarifa[c("LlaveTarifa")])), collapse=",")
          f1 <- str_replace(f1, "[(]", "")
          f1 <- str_replace(f1, "[c]", "")
          f1 <- str_replace(f1, "[)]", "")
          condiciones_consulta <- gsub("[\r\n]", "",paste0(condiciones_consulta,
                                                           " AND l.LlaveTarifa IN ( ",substr(f1,1,stop =1000000L)," ) ") )
        }
      }
    }
    
    if(!is.null(input$mercado)){
      if (seleccion_mercado != 'no') {
        ifelse (is.na(valores_validos_mercado),
                {
                  lista_seleccion <- input$mercado
                  valores_validos_mercado <<- valores_Mercado  %>%
                    filter(valores_Mercado$NombreMercado %in% lista_seleccion)
                } , NA) 
        if (nrow(valores_validos_mercado) == 0) {
          noHayValores <- TRUE
        } else {
          f1 <- t(unique(valores_validos_mercado[c("LlaveMercado")]))
          f1 <- paste0(apply(f1, 1, function(x) paste0(x)), collapse=",")
          #f1 <- paste0(gsub(":",",",unique(valores_validos_mercado[c("LlaveMercado")])), collapse=",")
          f1 <- str_replace(f1, "[(]", "")
          f1 <- str_replace(f1, "[c]", "")
          f1 <- str_replace(f1, "[)]", "")
          condiciones_consulta <- gsub("[\r\n]", "",paste0(condiciones_consulta,
                                                           " AND l.LlaveTipoObjeto IN ( ",substr(f1,1,stop =1000000L)," ) ") )
        }
      }
    }
    
    if(!is.null(input$pimt)){
      if (seleccion_PIMT != 'no') {
  
        if (nrow(valores_validos_PIMT) == 0) {
          noHayValores <- TRUE
        } else {
          f1 <- t(unique(valores_validos_PIMT[c("LlavePIMT")]))
          f1 <- paste0(apply(f1, 1, function(x) paste0(x)), collapse=",")
          #f1 <- paste0(gsub(":",",",unique(valores_validos_PIMT[c("LlavePIMT")])), collapse=",")
          f1 <- str_replace(f1, "[(]", "")
          f1 <- str_replace(f1, "[c]", "")
          f1 <- str_replace(f1, "[)]", "")
          condiciones_consulta <- gsub("[\r\n]", "",paste0(condiciones_consulta,
                                                           " AND l.LlavePIMT IN ( ",substr(f1,1,stop =1000000L)," ) ") )
        }
      }
    }
    
    if(!is.null(input$ynormalizacion)){
      if (seleccion_ynormalizacion != 'no') {
        if (nrow(valores_validos_ynormalizacion) == 0) {
          noHayValores <- TRUE
        } else {
          f1 <- t(unique(valores_validos_ynormalizacion[c("Llaveynormalizacion")]))
          f1 <- paste0(apply(f1, 1, function(x) paste0(x)), collapse=",")
          f1 <- paste0(gsub(":",",",unique(valores_validos_ynormalizacion[c("Llaveynormalizacion")])), collapse=",")
          f1 <- str_replace(f1, "[(]", "")
          f1 <- str_replace(f1, "[c]", "")
          f1 <- str_replace(f1, "[)]", "")
          condiciones_consulta <- gsub("[\r\n]", "",paste0(condiciones_consulta,
                                                           " AND Llaveynormalizacion IN ( ",substr(f1,1,stop =1000000L)," ) ") )
        }
      }
    }
    
    if(!is.null(input$habitacion)){
      if(input$habitacion == 'Habitado'){
        condiciones_consulta <- paste0(condiciones_consulta," AND l.Habitado = 1 ")
      }
      if(input$habitacion == 'No habitado'){      
        condiciones_consulta <- paste0(condiciones_consulta," AND l.Habitado = 0 ")
      }
      if(input$habitacion == 'Sin definir'){
        condiciones_consulta <- paste0(condiciones_consulta," AND l.Habitado is NULL ")
      }
    }

    # traer valores de Auditoria de lecturas

    if (1==0){#noHayValores) {
      showModal(modalDialog(
        title = tags$p(tags$span(tags$img(src="save.svg" ,alt="", width="24" ,height="24"),tags$strong("No hay resultados"),style="color: #0070ba"),style="text-align: center;"),
        
        "La consulta no genera resultados",
        footer = list(modalButton("Cancelar", icon = icon("fas fa-table"))
        ),
        easyClose =TRUE
      )
      )
    }  else {

      disable("ReiniciarControles")  
      disable("TraerDatos")
      
      condiciones_consultaGlobal <<- condiciones_consulta
      config <- config::get(config=configuracion.conexion)
      config.conexion <- config$conexion
      conexion <- odbcDriverConnect (config.conexion)
      
      cad.sql <- paste0(
          "select c.Tipo_generacion, c.Tipo_generacion AuxOrder, c.CodigoTipoCampaña, c.CodigoCampaña, c.NombreCampaña, c.NombreTipoCampaña,
          	SUM(CASE WHEN rc.CodigoresultadoCNR=5 THEN 1 ELSE 0 END) AS Pendientes,
          	SUM(CASE WHEN rc.CodigoresultadoCNR=1 THEN 1 ELSE 0 END) AS Normalizadas,
          	SUM(CASE WHEN rc.CodigoresultadoCNR=2 OR rc.CodigoresultadoCNR=1 THEN 1 ELSE 0 END) AS Cerradas,
          	SUM(CASE WHEN rc.CodigoresultadoCNR=3 THEN 1 ELSE 0 END) AS Fallidas,
          	SUM(CASE WHEN rc.CodigoresultadoCNR=4 THEN 1 ELSE 0 END) AS CNR,
	          CAST(REPLACE(max(vpi.ProcNorm), '%','') AS INT) PNorm, 
	          CAST(REPLACE(max(vpi.ProcCNR), '%','') AS INT) PCnr
          	from hecho.Inspeccion i
          		inner join Dimension.Campaña c on c.LlaveCampaña = i.LlaveCampaña
          		inner join Dimension.ResultadoCNR rc on rc.LlaveResultadoCNR = i.LlaveResultadoCNR
          		left join hecho.Llaves l on i.LlavePuntoConsumo = l.llavePuntoConsumo 
          		left join sgc.Excel.ValoresParaIconos vpi 
          			on c.CodigoCampaña = vpi.NombreGrupo COLLATE Latin1_General_100_CI_AS
          		where c.Tipo_generacion in ('Centralizada', 'Control de Calidad', 'Kronos', 'Regional') 
          			and c.Iniciativa in ('Masivos', 'Mantenimiento AP Masivos')
                and i.TaskTypeId in (15072, 10212, 15039, 10046, 10033)
                and (i.UnidadOperativa like 'NM%' OR i.UnidadOperativa is NULL)
                and ", condiciones_consulta,
          	"group by c.Tipo_generacion, c.CodigoTipoCampaña, c.CodigoCampaña, c.NombreCampaña, c.NombreTipoCampaña
             order by c.Tipo_generacion DESC"
          )
      
      datos.consulta <-sqlExecute(channel = conexion, query = cad.sql, fetch= T, as.is=T)
      
      odbcClose(conexion)
      
      if (nrow(datos.consulta ) > 0) {
        
        tabla.datos.valida <<- FALSE
        shinyjs::enable("Guardar")
        shinyjs::enable("Descargar")
        
        datos.consulta2 <- datos.consulta %>% group_by(Tipo_generacion) %>% 
          summarise(
            AuxOrder = 'zzzz',
            `CodigoTipoCampaña` = '', 
            `CodigoCampaña` = '', 
            `NombreCampaña` = 'Total', 
            NombreTipoCampaña = '',
            Pendientes = sum(Pendientes),
            Normalizadas = sum(Normalizadas),
            Cerradas = sum(Cerradas),
            Fallidas = sum(Fallidas),
            CNR = sum(CNR),
            PNorm = NA,
            PCnr = NA,
            .groups = 'drop') %>%
          as.data.frame()
        datos.consulta <- union(datos.consulta,datos.consulta2);
        
        datos.consulta$Bolsa <- datos.consulta$Pendientes + datos.consulta$Cerradas + datos.consulta$Fallidas
        datos.consulta$PrcAvance <- trunc((datos.consulta$Cerradas + datos.consulta$Fallidas)/datos.consulta$Bolsa*10000)/100
        datos.consulta$PrcNormalizacion <- trunc(datos.consulta$Normalizadas/datos.consulta$Cerradas*10000)/100
        datos.consulta$PrcCNR <- trunc(datos.consulta$CNR/datos.consulta$Cerradas*10000)/100
        datos.consulta$PrcFallidas <- trunc(datos.consulta$Fallidas/datos.consulta$Bolsa*10000)/100
        
        datos.consulta$PrcAvanceText <- paste0(
          ifelse(is.infinite(datos.consulta$PrcAvance) | is.nan(datos.consulta$PrcAvance),"---",datos.consulta$PrcAvance)
          ,"%")
        datos.consulta$PrcFallidasText <- paste0(
          ifelse(is.infinite(datos.consulta$PrcFallidas) | is.nan(datos.consulta$PrcFallidas),"---",datos.consulta$PrcFallidas)
          ,"%")
        
        datos.consulta$PrcNormalizacionIcon <- case_when(
          is.na(datos.consulta$PNorm) | is.infinite(datos.consulta$PrcNormalizacion) | is.nan(datos.consulta$PrcNormalizacion) ~ "",
          datos.consulta$PrcNormalizacion < datos.consulta$PNorm-1 ~ "<i class='fas fa-arrow-alt-circle-down' style='color:#ff0000'></i>",
          datos.consulta$PrcNormalizacion >= datos.consulta$PNorm-1 & datos.consulta$PrcNormalizacion < datos.consulta$PNorm+1  ~ "<i class='fas fa-arrow-alt-circle-left' style='color:#ff9900'></i>",
          datos.consulta$PrcNormalizacion >= datos.consulta$PNorm+1 ~ "<i class='fas fa-arrow-alt-circle-up' style='color:#009933'></i>"
        )
        datos.consulta$PrcNormalizacionText <- paste0(
          ifelse(is.infinite(datos.consulta$PrcNormalizacion) | is.nan(datos.consulta$PrcNormalizacion),"---",datos.consulta$PrcNormalizacion)
          ,"% "
          ,datos.consulta$PrcNormalizacionIcon)
        
        datos.consulta$PrcCNRIcon <- case_when(
          is.na(datos.consulta$PCnr) | is.infinite(datos.consulta$PrcCNR) | is.nan(datos.consulta$PrcCNR) ~ "",
          datos.consulta$PrcCNR < datos.consulta$PCnr-.6 ~ "<i class='fas fa-arrow-alt-circle-down' style='color:#ff0000'></i>",
          datos.consulta$PrcCNR >= datos.consulta$PCnr-.6 & datos.consulta$PrcCNR < datos.consulta$PCnr+.6  ~ "<i class='fas fa-arrow-alt-circle-left' style='color:#ff9900'></i>",
          datos.consulta$PrcCNR >= datos.consulta$PCnr+.6 ~ "<i class='fas fa-arrow-alt-circle-up' style='color:#009933'></i>"
        )
        datos.consulta$PrcCNRText <- paste0(
          ifelse(is.infinite(datos.consulta$PrcCNR) | is.nan(datos.consulta$PrcCNR),"---",datos.consulta$PrcCNR)
          ,"% "
          ,datos.consulta$PrcCNRIcon)
        
        
        datos.consulta <<- datos.consulta
        
        tabla_despliegue();
        
      } else{
        disable("Guardar")
        disable("Descargar")
        showModal(modalDialog(
          title = "La consulta no encuentra datos que cumplan.",
          footer = tags$div(id="modal1",modalButton("Cerrar")),
          easyClose = TRUE
        ))
      }
      shinyjs::enable("ReiniciarControles")
      shinyjs::enable("TraerDatos")
    }

  }, ignoreInit = TRUE)
  
  data.grafica<-reactive({
    if (exists("tabla.datos")){
      tabla.datos.valida <<- FALSE
      tmp<-tabla.datos
      sel<-tryCatch(tabla.datos[(selected()$pointNumber+1),,drop=FALSE] , error=function(e){NULL})
      if (nrow(sel) > 0 & tipo.grafica == 'B') {
        secuencia.sel <- selected()$pointNumber + 1
        valores.ceros <- sort(unique(tabla.datos$AcumuladoConsumoCero))[secuencia.sel]
        sel <-  tabla.datos %>% filter(AcumuladoConsumoCero %in% valores.ceros)
        list(data=tmp,sel=sel)
      } else {
        click('MostrarTabla')
        list(data=tmp,sel=sel)
      }
      
    }
    
   # list(data=tmp,sel=sel)
    
  })
 
  # Presentación del mapa --------------------------------------------------- 
  output$map <- renderGoogle_map({
    if (exists("tabla.datos")) {
      mapa.despliegue(tabla.datos, data.grafica()$sel, datos.consulta, FALSE, tabla.datos.valida, input) #, datos.imagen, input)
    }
  })
  
  # Cambiar tipo de gráfica  ----------------------------------------
  observeEvent(input$selec.graf,{

    valorSwitch <- input$selec.graf
    if (valorSwitch == TRUE) {
      tipo.grafica <<- 'B'
      output$plot <- renderPlotly({
        p <- ggplot(tabla.datos, mapping = aes(x = AcumuladoConsumoCero)) +
          geom_bar(color ="#0070ba",fill= "#0070ba",size =1) + theme_bw()+
          ylab("Conteo") +
          xlab("Períodos consecutivos en cero")
        # Manejo de selección en gráfica
        obj <- data.grafica()$sel
        if(nrow(obj)!=0) {
          p <- p + geom_bar(data=obj,color="orange",fill="orange")
        }
        ggplotly(p,source="master")
      })
    } else {
      tipo.grafica <<- 'P'
      output$plot <- renderPlotly({
        p <-
          ggplot(tabla.datos,
                 mapping = aes(x = AcumuladoConsumoCero, y = PromedioActiva6CN, text = texto)) +
          geom_point(color = "#0070ba", size = 1) + theme_bw() +
          ylab("Consumo promedio 6 meses anteriores [kWh]") +
          xlab("Períodos consecutivos en cero")
        
        # Manejo de selección en gráfica
        obj <- data.grafica()$sel
        if(nrow(obj)!=0) {
          p <- p + geom_point(data=obj,color="orange")
        }
        ggplotly(p, source = "master")
      })
    }
    }, ignoreInit = TRUE)
  
  # Seleccionar o deselecionar todos ----------------------------------------
  
  observeEvent(input$selec.tod,{
    tabla.temp <<- hot_to_r(input$datos)    # Tabla temporal para alimentar rhandsontable
    
    if(input$selec.tod == "Todos"){
      tabla.temp[,1] <- TRUE             # Cambia primera columna a seleccionado
    }
    
    if(input$selec.tod == "Ninguno"){
      tabla.temp[,1] <- FALSE               # Cambia primera columna a deseleccionado
    }
    # output$datos <- renderRHandsontable({   # Actualiza rhandsontable
    #   tabla.despliegue(tabla.temp, data.grafica()$sel)
    # })
    # output$map <- renderGoogle_map({
    #   mapa.despliegue(tabla.temp, data.grafica()$sel, datos.consulta, TRUE,tabla.datos.valida, input) #, datos.imagen, input)
    # })
    
  }, ignoreInit = TRUE)
  
  # Descargar excel -----------------------------------------------------------
  
  output$Descargar <- downloadHandler(
    
    filename = function() { 
      #paste0("Ceros", Sys.Date(), ".csv")
      paste0("Ceros", Sys.Date(), ".xlsx")
    },
    content = function(file) {
      
      ordenes <- transf.rhand(NULL)                      # Transforma los datos tomados del objeto rhandsontable.
      #ordenes <- ordenes, ordenes[,1] == TRUE)[,-1]      # Filtra aquellos con checkbox activo y omite primera columna
      wbk <- createWorkbook()
      addWorksheet( wb = wbk,  sheetName = "Data" )
      setColWidths(wbk,1,cols = 1:2,
                   widths = c(6,10))
      # Ver https://www.r-bloggers.com/2019/07/excel-report-generation-with-shiny/
      # writeData(wbk,sheet = 1,
      #           c("CodigoPuntoConsumo" ,"NombreCliente","DireccionPuntoConsumo","Seccion","AcumuladoConsumoCero"         ,"PromedioActiva6CN"   ,"NombreZona" ,"NombreRegion","NombreOficina"
      #             ,"NombreDepartamentoGeografia","NombreMunicipioGeografia","NombreLocalidadZonaGeografia","Nombre_corto_SMT"),
      #           startRow = 1,startCol = 1)
      
      addStyle(wbk,sheet = 1, style = createStyle( fontSize = 12, textDecoration = "bold" ), rows = 1:1, cols = 1:ncol(ordenes) )
      
      writeData(wbk,sheet = 1,ordenes,startRow = 1,startCol = 1 )
      saveWorkbook(wbk, file)
      #write.table(ordenes,file,sep=separador.csv, fileEncoding = "latin1")
    }#,
    #contentType = "csv"
  )
  
  observeEvent(input$MostrarTabla, {
    if (exists("tabla.datos")) {
      if (nrow(data.grafica()$sel) == 0) {
        tabla.temp <- tabla.datos
      }
      if (exists("tabla.temp")){
        tabla.temp[,1] <- TRUE
        # output$datos <- renderRHandsontable({   # Actualiza rhandsontable
        #   tabla.despliegue(tabla.temp, data.grafica()$sel)
        # })
      } 
    }
  }, ignoreInit = TRUE)  
  
  observeEvent(input$MostrarMapa, {
    if ( exists("tabla.datos")) {
      mapa.despliegue(tabla.datos, data.grafica()$sel, datos.consulta,TRUE,tabla.datos.valida, input) #, datos.imagen, input)
    }
    #  mapa.despliegue(tabla.datos, data.grafica()$sel, datos.consulta,TRUE,tabla.datos.valida, input)
    # output$map <- renderGoogle_map({
    #   mapa.despliegue(tabla.datos, data.grafica()$sel, datos.consulta)
    # })
  }, ignoreInit = TRUE)
  
  # Generar órdenes de inspección en BD ---------------------------------------------------
  observeEvent(input$Guardar,{
 
    showModal(modalDialog(
      
      title = tags$p(tags$span(tags$img(src="save.svg" ,alt="", width="24" ,height="24"),tags$strong("Generar órdenes"),style="color: #0070ba"),style="text-align: center;"),
      
      "Se guardarán los clientes seleccionados para ser incluidos \n en la lista de órdenes de inspección",
      hr(),
      
      selectInput("variable", 
                  "Seleccione campaña:",
                  nom.camp),
      textAreaInput("observ.orden", label = "Observaciones:",
                    height = "100px", rows = 4, 
                    placeholder = "Comentarios para acompañar la órden", 
                    resize = "vertical"),
      # footer = list(tags$div(id="modal1",modalButton("Cancelar", icon = icon("fas fa-table"))),
      #               actionButton("Guardar2", "Guardar",icon = icon("fas fa-table"),
      #                            style="color: #fff; background-color: #0070ba; border-color: #0070ba")
      # ),
      footer = fluidRow(column(width=2, offset=7,tags$div(id="modal1", 
                                                          modalButton("Cancelar", icon = icon("fas fa-table")))),
                        column(width=2,actionButton("Guardar2", "Guardar",icon = icon("fas fa-table"),
                                                    style="color: #fff; background-color: #0070ba; border-color: #0070ba"))
      ),
      easyClose =TRUE
    )
    )
  }, ignoreInit = TRUE)
  
  
  observeEvent(input$Guardar2, {                                          # Código para botón diálogo modal
    
    #  Grabar órdenes en BD
    
    ordenes <- transf.rhand(input$datos)                      # Transforma los datos tomados del objeto rhandsontable.
    ordenes <- subset(ordenes, ordenes[,1] == TRUE)[,-1]      # Filtra aquellos con checkbox activo y omite primera columna
    
    if (nrow(ordenes) > 0) {
      ordenes<-cbind(ordenes[,1],                                         # Punto de consumo: 'CodigoPuntoConsumo'
                     rep(as.character(input$variable), nrow(ordenes)),    # Línea para tipo de campaña seleccionada: 'NombreCampana'
                     rep("Calidad de Prediccion", nrow(ordenes)),          # Reporte origen de las órdenes: 'ReporteOrigen'
                     rep(as.character(Sys.time()), nrow(ordenes)),        # Fecha y hora de generación de las órdenes: 'FechaGeneracion'
                     rep(analista, nrow(ordenes)),                        # Usuario que genera las órdenes: 'AnalistaSolicitante'
                     rep(0.5, nrow(ordenes)),                             # 'Probabilidad' de detección. ** Se debe incluir probabilidad de la campaña, por ahora es 0.5
                     ordenes[,9],                                         # 'RecuperacionMedia': Se asume promedio anterior
                     rep(as.character(input$observ.orden), nrow(ordenes)) #  Observación que acompaña la orden 'ObservacionOrden'
      )
      
      ordenes<-data.frame(ordenes)
      ordenes[,7]<-as.numeric(as.character(ordenes[,7]))                  # Se convierte a texto el dato tomado de la tabla
      
      #conexion2 <- odbcDriverConnect ("driver={SQL Server};server=WMENERTOLIMA\\SQL2017;database=DWSRBI_KRONOS;Uid=profesional.bi;Pwd=123Canea7;trustedconnection=true" )
      config <- config::get(config=configuracion.conexion)
      config.conexion <- config$conexion
      conexion <- odbcDriverConnect (config.conexion)
      cad.sql<-" INSERT INTO [DWSRBI_KRONOS].[Hecho].[OrdenInspeccion]([CodigoPuntoConsumo],[NombreCampana],[ReporteOrigen],[FechaGeneracion],[AnalistaSolicitante],[Probabilidad],[RecuperacionMedia],[ObservacionOrden])"
      cad.sql<-paste (cad.sql,"VALUES (?,?,?,?,?,?,?,?)")
      
      sqlExecute(channel = conexion,
                 cad.sql,
                 data = ordenes)              # En 'ordenes' se incluyen las órdenes a ser incluidas en la tabla 'OrdenInspeccion'
      odbcClose(conexion)
      
      ord.gen <<- ord.gen + nrow(ordenes)     # Se actualiza el contador de órdenes generadas en la sesión. Variable global
      
      output$MensajeOrdenes <- renderMenu({   # Actualiza mensaje en la barra de menú para órdenes generadas
        
        dropdownMenu(type = "notifications",  # Mensaje en barra de encabezado para indicar generación de órdenes
                     headerText = "Tiene una notificación",
                     icon =icon("fas fa-comments"),
                     notificationItem(
                       text = paste(ord.gen, "órdenes generadas en esta sesión"),
                       icon("fas fa-gavel")
                     )
        )
      })
    }
    removeModal()
  })
  
})



# Ejecutar aplicación -----------------------------------------------------


shinyApp(ui = ui, server = server)