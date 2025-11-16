library(shiny)
library(shinydashboard)
library(haven)
library(readxl)
library(foreign)
library(psych)
library(GPArotation)
library(ggplot2)
library(shinythemes)
library(DT)
library(readr)
library(semTools)
library(plotly)
library(officer)
library(flextable)
library(lavaan)
library(semPlot)
library(dplyr)
library(shinyjs)
library(grid)
library(gridExtra)
library(pander)

ui <- dashboardPage(
  dashboardHeader(
    title = span(
      "FAnalyzr",
      style = "font-family: 'Poppins', sans-serif; font-weight: 700; font-size: 1.8rem; 
               color: #000000 !important; text-shadow: 0 2px 4px rgba(0,0,0,0.1);"
    ),
    titleWidth = 250
  ),
  dashboardSidebar(
    width = 300,
    tags$head(
      tags$style(HTML("
        @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&family=Poppins:wght@400;500;600;700&display=swap');
        
        :root {
          --primary-blue: #3498db;
          --primary-purple: #9b59b6;
          --primary-green: #2ecc71;
          --dark-blue: #2980b9;
          --dark-purple: #8e44ad;
          --dark-green: #27ae60;
          --dark-gray: #2c3e50;
          --medium-gray: #34495e;
          --light-gray: #ecf0f1;
          --background-gray: #f8f9fa;
          --border-gray: #e0e0e0;
          --text-dark: #2c3e50;
          --text-medium: #7f8c8d;
          --text-light: #95a5a6;
          --success-light: #d5f4e6;
          --info-light: #e8f4fc;
          --warning-light: #fff3cd;
        }
        
        /* Sidebar styling */
        .sidebar {
          background: linear-gradient(180deg, #ffffff 0%, #f8f9fa 100%) !important;
          border-right: 1px solid var(--border-gray);
          box-shadow: 2px 0 10px rgba(0,0,0,0.1);
        }
        
        .sidebar-menu {
          background: transparent !important;
        }
        
        .sidebar-menu li a {
          font-family: 'Roboto', sans-serif;
          font-weight: 500;
          color: var(--text-dark) !important;
          border-left: 4px solid transparent;
          transition: all 0.3s ease;
          margin: 5px 10px;
          border-radius: 8px;
        }
        
        .sidebar-menu li a:hover {
          background: linear-gradient(135deg, var(--info-light) 0%, #e3f2fd 100%) !important;
          color: var(--primary-blue) !important;
          border-left: 4px solid var(--primary-blue);
        }
        
        .sidebar-menu li.active a {
          background: linear-gradient(135deg, var(--primary-blue) 0%, var(--primary-purple) 100%) !important;
          color: white !important;
          border-left: 4px solid var(--primary-blue);
          box-shadow: 0 4px 12px rgba(52, 152, 219, 0.3);
        }
        
        /* Main content styling */
        .content-wrapper, .right-side {
          background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%) !important;
          background-attachment: fixed;
          height: 100vh;
          overflow: hidden;
        }
        
        /* Main content area - fixed height with scrollable results */
        .main-content-container {
          height: calc(100vh - 80px);
          display: flex;
          flex-direction: column;
        }
        
        .fixed-sidebar {
          height: calc(100vh - 80px);
          overflow-y: auto;
        }
        
        .scrollable-results {
          height: calc(100vh - 80px);
          overflow-y: auto;
          background: linear-gradient(180deg, #ffffff 0%, #fafbfc 100%);
          border-radius: 15px;
          box-shadow: 0 8px 25px rgba(0, 0, 0, 0.1);
          padding: 25px;
          border: 1px solid rgba(255,255,255,0.2);
        }
        
        .box {
          background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%);
          border-radius: 15px;
          border: none;
          box-shadow: 0 8px 25px rgba(0, 0, 0, 0.1);
          margin-bottom: 20px;
          border: 1px solid rgba(255,255,255,0.2);
        }
        
        /* Tab styling */
        .nav-tabs-custom {
          background: transparent;
          border: none;
          box-shadow: none;
        }
        
        .nav-tabs-custom > .nav-tabs {
          background: transparent;
          border-bottom: 2px solid var(--border-gray);
          margin-bottom: 20px;
        }
        
        .nav-tabs-custom > .nav-tabs > li.active > a {
          background: linear-gradient(135deg, var(--primary-blue) 0%, var(--primary-purple) 100%);
          color: white;
          border: none;
          border-radius: 10px 10px 0 0;
          box-shadow: 0 -4px 12px rgba(52, 152, 219, 0.3);
          font-weight: 600;
        }
        
        .nav-tabs-custom > .nav-tabs > li > a {
          color: var(--text-medium);
          border: none;
          border-radius: 10px 10px 0 0;
          margin-right: 5px;
          transition: all 0.3s ease;
          font-weight: 500;
        }
        
        .nav-tabs-custom > .nav-tabs > li > a:hover {
          background: linear-gradient(135deg, var(--light-gray) 0%, #e9ecef 100%);
          color: var(--primary-blue);
        }
        
        .tab-content {
          background: transparent;
          border-radius: 0 15px 15px 15px;
          padding: 0;
          border: none;
        }
        
        /* Button styling */
        .btn-primary {
          background: linear-gradient(135deg, var(--primary-blue) 0%, var(--dark-blue) 100%);
          border: none;
          color: white;
          font-weight: 500;
          border-radius: 8px;
          padding: 10px 20px;
          transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
          box-shadow: 0 4px 12px rgba(52, 152, 219, 0.3);
        }
        
        .btn-primary:hover {
          background: linear-gradient(135deg, var(--dark-blue) 0%, #2472a4 100%);
          transform: translateY(-2px);
          box-shadow: 0 6px 20px rgba(52, 152, 219, 0.4);
          border: none;
        }
        
        .btn-success {
          background: linear-gradient(135deg, var(--primary-green) 0%, var(--dark-green) 100%);
          border: none;
          border-radius: 8px;
          transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
          box-shadow: 0 4px 12px rgba(46, 204, 113, 0.3);
        }
        
        .btn-success:hover {
          background: linear-gradient(135deg, var(--dark-green) 0%, #219653 100%);
          transform: translateY(-2px);
          box-shadow: 0 6px 20px rgba(46, 204, 113, 0.4);
          border: none;
        }
        
        .btn-info {
          background: linear-gradient(135deg, var(--primary-purple) 0%, var(--dark-purple) 100%);
          border: none;
          border-radius: 8px;
          transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
          box-shadow: 0 4px 12px rgba(155, 89, 182, 0.3);
        }
        
        .btn-info:hover {
          background: linear-gradient(135deg, var(--dark-purple) 0%, #7d3c98 100%);
          transform: translateY(-2px);
          box-shadow: 0 6px 20px rgba(155, 89, 182, 0.4);
          border: none;
        }
        
        /* Input styling */
        .form-control, .selectize-input {
          border-radius: 8px !important;
          border: 1px solid var(--border-gray) !important;
          transition: all 0.3s ease;
          padding: 10px 12px;
        }
        
        .form-control:focus, .selectize-input.focus {
          border-color: var(--primary-blue) !important;
          box-shadow: 0 0 0 3px rgba(52, 152, 219, 0.1) !important;
          outline: none;
        }
        
        /* Well panel styling */
        .well {
          background: linear-gradient(135deg, #ffffff 0%, var(--light-gray) 100%);
          border-radius: 12px;
          border: 1px solid var(--border-gray);
          box-shadow: 0 4px 15px rgba(0, 0, 0, 0.08);
          transition: all 0.3s ease;
        }
        
        .well:hover {
          box-shadow: 0 6px 20px rgba(0, 0, 0, 0.12);
        }
        
        /* File input styling */
        .file-input {
          background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
          border-radius: 8px;
          border: 2px dashed #bdc3c7;
          transition: all 0.3s ease;
          padding: 15px;
        }
        
        .file-input:hover {
          border-color: var(--primary-blue);
          background: linear-gradient(135deg, #e3f2fd 0%, #f8f9fa 100%);
        }
        
        /* Syntax help boxes */
        .syntax-help {
          background: linear-gradient(135deg, var(--info-light) 0%, #e3f2fd 100%);
          border-left: 4px solid var(--primary-blue);
          padding: 15px;
          margin: 15px 0;
          border-radius: 8px;
          box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
        }
        
        /* Instruction boxes */
        .instruction-box {
          background: linear-gradient(135deg, var(--info-light) 0%, #e3f2fd 100%);
          border-left: 4px solid var(--primary-blue);
          padding: 20px;
          margin-bottom: 25px;
          border-radius: 10px;
          box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
        }
        
        /* Interpretation boxes */
        .interpretation-box {
          background: linear-gradient(135deg, var(--success-light) 0%, #e8f5e8 100%);
          border-left: 4px solid var(--primary-green);
          padding: 20px;
          margin: 25px 0;
          border-radius: 10px;
          box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
        }
        
        /* Headers */
        h1, h2, h3, h4 {
          font-family: 'Poppins', sans-serif;
          font-weight: 600;
          color: var(--text-dark);
        }
        
        h3 {
          background: linear-gradient(135deg, var(--primary-blue) 0%, var(--primary-purple) 100%);
          -webkit-background-clip: text;
          -webkit-text-fill-color: transparent;
          background-clip: text;
          padding-bottom: 10px;
          margin-bottom: 20px;
          border-bottom: 2px solid var(--border-gray);
        }
        
        /* DataTables styling */
        .dataTables_wrapper {
          background: white;
          border-radius: 10px;
          box-shadow: 0 4px 12px rgba(0,0,0,0.05);
          padding: 15px;
        }
        
        /* Custom scrollbar */
        ::-webkit-scrollbar {
          width: 8px;
        }
        
        ::-webkit-scrollbar-track {
          background: #f1f1f1;
          border-radius: 4px;
        }
        
        ::-webkit-scrollbar-thumb {
          background: linear-gradient(135deg, var(--primary-blue) 0%, var(--primary-purple) 100%);
          border-radius: 4px;
        }
        
        ::-webkit-scrollbar-thumb:hover {
          background: linear-gradient(135deg, var(--dark-blue) 0%, var(--dark-purple) 100%);
        }
        
        /* Footer styling */
        .footer {
          text-align: center;
          padding: 20px;
          margin-top: 30px;
          color: var(--text-medium);
          font-size: 14px;
          border-top: 1px solid var(--border-gray);
        }
        
        /* Analysis step styling */
        .analysis-step {
          background: linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%);
          border-radius: 10px;
          padding: 20px;
          margin-bottom: 20px;
          border-left: 4px solid var(--primary-purple);
          box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
          transition: all 0.3s ease;
        }
        
        .analysis-step:hover {
          transform: translateX(5px);
          box-shadow: 0 6px 18px rgba(0, 0, 0, 0.08);
        }
        
        /* Diagram controls */
        .diagram-controls {
          background: linear-gradient(180deg, #ffffff 0%, #f8f9fa 100%);
          border-radius: 12px;
          box-shadow: 0 4px 15px rgba(0, 0, 0, 0.08);
          padding: 20px;
          margin-bottom: 20px;
          border: 1px solid var(--border-gray);
        }
        
        /* Loading animation */
        .shiny-notification {
          background: linear-gradient(135deg, var(--primary-blue) 0%, var(--primary-purple) 100%);
          color: white;
          border-radius: 10px;
          box-shadow: 0 8px 25px rgba(0,0,0,0.15);
          border: none;
        }
        
        /* Results content styling */
        .results-content {
          max-height: 100%;
        }
        
        .results-section {
          margin-bottom: 30px;
          padding: 20px;
          background: white;
          border-radius: 10px;
          box-shadow: 0 4px 12px rgba(0,0,0,0.05);
          border: 1px solid var(--border-gray);
        }
      "))
    ),
    sidebarMenu(
      id = "tabs",
      menuItem("Exploratory Factor Analysis", tabName = "efa", icon = icon("search"), selected = TRUE),
      menuItem("Confirmatory Factor Analysis", tabName = "cfa", icon = icon("check-circle")),
      menuItem("Regression Table", tabName = "regression", icon = icon("table"))
    )
  ),
  dashboardBody(
    useShinyjs(),
    div(class = "main-content-container",
        tabItems(
          # EFA Tab
          tabItem(tabName = "efa",
                  fluidRow(
                    column(12,
                           div(class = "box",
                               h2("Exploratory Factor Analysis", style = "text-align: center; margin-bottom: 30px;")
                           )
                    )
                  ),
                  fluidRow(
                    column(4,
                           div(class = "fixed-sidebar",
                               div(class = "box",
                                   h4("Data Input & Parameters", icon("database")),
                                   fileInput("file_efa", "Upload Dataset (CSV, Excel, SAV, DTA)", 
                                             accept = c(".csv", ".xlsx", ".xls", ".sav", ".dta"),
                                             width = "100%"),
                                   uiOutput("varSelect_efa"),
                                   numericInput("nfactors", "Number of Factors to Extract:", value = 2, min = 1),
                                   selectInput("fm", "Extraction Method:", 
                                               choices = c("minres", "ml", "pa", "wls", "gls", "uls", "principal")),
                                   selectInput("rotate", "Rotation Method:", 
                                               choices = c("none", "varimax", "promax", "oblimin", "simplimax", "quartimin", "geominQ")),
                                   selectInput("reliability", "Select Reliability Method:", 
                                               choices = c("Cronbach Alpha", "McDonald Omega")),
                                   actionButton("analyze", "Run EFA", class = "btn-primary", icon = icon("play")),
                                   br(), br(),
                                   downloadButton("downloadAllEFA", "Download All EFA Results (PDF)", class = "btn-success")
                               ),
                               div(class = "syntax-help",
                                   h5("EFA Analysis Notes:", icon("info-circle")),
                                   tags$ul(
                                     tags$li("Select variables for factor analysis"),
                                     tags$li("Choose number of factors based on scree plot and parallel analysis"),
                                     tags$li("KMO > 0.6 and Bartlett's p < 0.05 indicate suitable data"),
                                     tags$li("Factor loadings > |0.4| are typically considered meaningful")
                                   )
                               )
                           )
                    ),
                    column(8,
                           div(class = "scrollable-results",
                               tabsetPanel(
                                 tabPanel("Scree Plot", 
                                          plotlyOutput("screePlot"), 
                                          br(), 
                                          downloadButton("downloadScree", "Download Scree Plot", class = "btn-info")
                                 ),
                                 tabPanel("Parallel Analysis", 
                                          verbatimTextOutput("parallelInterpretation"),
                                          plotOutput("faParallelPlot"), 
                                          br(), 
                                          downloadButton("downloadParallel", "Download Parallel Plot", class = "btn-info")
                                 ),
                                 tabPanel("KMO & Bartlett's Test", verbatimTextOutput("kmoBartlett")),
                                 tabPanel("Variance Explained", verbatimTextOutput("varianceTable")),
                                 tabPanel("EFA Results", 
                                          verbatimTextOutput("efaResults"), 
                                          br(),
                                          downloadButton("downloadEFAWord", "Download EFA Results (Word)", class = "btn-info")
                                 ),
                                 tabPanel("Reliability", 
                                          verbatimTextOutput("reliabilityResult"), 
                                          downloadButton("downloadDiagram", "Download Factor Diagram", class = "btn-info")
                                 ),
                                 tabPanel("Correlation Matrix", 
                                          DTOutput("corMatrix"),
                                          tags$div(
                                            style = "margin-top: 10px; font-style: italic; font-size: 14px; color: var(--text-medium);",
                                            "*p < 0.05, **p < 0.01, ***p < 0.001"
                                          )
                                 )
                               )
                           )
                    )
                  )
          ),
          
          # CFA Tab
          tabItem(tabName = "cfa",
                  fluidRow(
                    column(12,
                           div(class = "box",
                               h2("Confirmatory Factor Analysis", style = "text-align: center; margin-bottom: 30px;")
                           )
                    )
                  ),
                  fluidRow(
                    column(4,
                           div(class = "fixed-sidebar",
                               div(class = "box",
                                   h4("Data Input & Model Specification", icon("database")),
                                   fileInput("file_cfa", "Upload Dataset (CSV, Excel, SAV, DTA)", 
                                             accept = c(".csv", ".xlsx", ".xls", ".sav", ".dta"),
                                             width = "100%"),
                                   uiOutput("varSelect_cfa"),
                                   selectInput("estimator", "Select Estimator:",
                                               choices = c("ML" = "ML", "MLR" = "MLR", "WLSMV" = "WLSMV"),
                                               selected = "MLR"),
                                   actionButton("runModel", "Run CFA", class = "btn-success", icon = icon("play")),
                                   br(), br(),
                                   downloadButton("downloadAllCFA", "Download All CFA Results (PDF)", class = "btn-success")
                               ),
                               div(class = "syntax-help",
                                   h5("CFA lavaan Syntax Examples:", icon("code")),
                                   tags$code("
        # Basic CFA model:
        f1 =~ item1 + item2 + item3
        f2 =~ item4 + item5 + item6
        
        # With factor correlations:
        f1 ~~ f2
        
        # With error covariances:
        item1 ~~ item2
                        "),
                                   tags$ul(
                                     style = "margin-top: 10px;",
                                     tags$li("=~ defines latent variables"),
                                     tags$li("~~ defines covariances/correlations"),
                                     tags$li("~ defines regressions")
                                   )
                               )
                           )
                    ),
                    column(8,
                           div(class = "scrollable-results",
                               div(class = "box",
                                   h4("Model Specification (lavaan syntax)", icon("edit")),
                                   textAreaInput("modelText", NULL, height = "200px", width = "100%",
                                                 placeholder = "Enter your CFA model using lavaan syntax...\n\nExample:\nf1 =~ x1 + x2 + x3\nf2 =~ x4 + x5 + x6\nf1 ~~ f2"),
                                   tabsetPanel(
                                     tabPanel("Model Summary", verbatimTextOutput("modelSummary")),
                                     tabPanel("Standardized Estimates", verbatimTextOutput("standardized")),
                                     tabPanel("Unstandardized Estimates", verbatimTextOutput("unstandardized")),
                                     tabPanel("Modification Indices", DTOutput("modIndices")),
                                     tabPanel("Reliability & Validity", verbatimTextOutput("reliabilityResults")),
                                     tabPanel("SEM Plot", 
                                              div(class = "diagram-controls",
                                                  fluidRow(
                                                    column(6,
                                                           textInput("diagram_title_cfa", "Diagram Title:", 
                                                                     placeholder = "Enter diagram title"),
                                                           selectInput("diagram_layout_cfa", "Layout:",
                                                                       choices = c("tree", "circle", "spring", "tree2", "tree3"),
                                                                       selected = "tree")
                                                    ),
                                                    column(6,
                                                           downloadButton("downloadPlot", "Download Plot (PNG)", class = "btn-info"),
                                                           br(), br(),
                                                           sliderInput("node_size_cfa", "Node Size:", 
                                                                       min = 5, max = 20, value = 10, step = 1)
                                                    )
                                                  )
                                              ),
                                              plotOutput("semPlot", height = "600px")
                                     ),
                                     tabPanel("Residual Covariances", verbatimTextOutput("residCov")),
                                     tabPanel("Discriminant Validity", tableOutput("discriminantValidity"))
                                   )
                               )
                           )
                    )
                  )
          ),
          
          # Regression Table Tab
          tabItem(tabName = "regression",
                  fluidRow(
                    column(12,
                           div(class = "box",
                               h2("Regression Analysis", style = "text-align: center; margin-bottom: 30px;")
                           )
                    )
                  ),
                  fluidRow(
                    column(4,
                           div(class = "fixed-sidebar",
                               div(class = "box",
                                   h4("Data Input & Model Specification", icon("database")),
                                   fileInput("file_reg", "Upload Dataset (CSV, Excel, SAV, DTA)", 
                                             accept = c(".csv", ".xlsx", ".xls", ".sav", ".dta"),
                                             width = "100%"),
                                   uiOutput("varSelect_reg"),
                                   selectInput("regEstimator", "Select Estimator:",
                                               choices = c("ML" = "ML", "MLR" = "MLR", "WLSMV" = "WLSMV"),
                                               selected = "MLR"),
                                   actionButton("runRegModel", "Run Regression", class = "btn-info", icon = icon("play")),
                                   br(), br(),
                                   downloadButton("downloadAllReg", "Download All Regression Results (PDF)", class = "btn-success")
                               ),
                               div(class = "syntax-help",
                                   h5("Regression lavaan Syntax Examples:", icon("code")),
                                   tags$code("
        # Simple regression:
        y ~ x1 + x2 + x3
        
        # Multiple outcomes:
        y1 ~ x1 + x2
        y2 ~ x1 + x3
        
        # With interactions:
        y ~ x1 + x2 + x1:x2
                        "),
                                   tags$ul(
                                     style = "margin-top: 10px;",
                                     tags$li("~ defines regression relationships"),
                                     tags$li(": creates interaction terms"),
                                     tags$li("Multiple equations supported")
                                   )
                               )
                           )
                    ),
                    column(8,
                           div(class = "scrollable-results",
                               div(class = "box",
                                   h4("Regression Model (lavaan syntax)", icon("edit")),
                                   textAreaInput("regModelText", NULL, height = "150px", width = "100%",
                                                 placeholder = "Enter your regression model using lavaan syntax...\n\nExample:\ny ~ x1 + x2 + x3 + control1 + control2"),
                                   DTOutput("regressionTable")
                               )
                           )
                    )
                  )
          )
        )
    ),
    tags$footer(
      class = "footer",
      div(
        "For suggestions or assistance with funding a Heroku plan, please contact me:",
        br(),
        tags$a(href = "mailto:mudassiribrahim30@gmail.com", "mudassiribrahim30@gmail.com", 
               style = "color: var(--primary-blue);")
      )
    )
  )
)

# Server function with robust PDF download functionality
server <- function(input, output, session) {
  # Shared data reactive
  data_shared <- reactive({
    # Check which file input was used
    if (!is.null(input$file_efa)) {
      file <- input$file_efa
    } else if (!is.null(input$file_cfa)) {
      file <- input$file_cfa
    } else if (!is.null(input$file_reg)) {
      file <- input$file_reg
    } else {
      return(NULL)
    }
    
    req(file)
    ext <- tools::file_ext(file$name)
    switch(ext,
           csv = read_csv(file$datapath),
           xlsx = read_excel(file$datapath),
           xls = read_excel(file$datapath),
           sav = read_sav(file$datapath),
           dta = read_dta(file$datapath),
           validate("Unsupported file type")
    )
  })
  
  # Consistent variable selection UI for all tabs
  output$varSelect_efa <- renderUI({
    req(data_shared())
    selectInput("vars_efa", "Select Variables:", choices = names(data_shared()), multiple = TRUE)
  })
  
  output$varSelect_cfa <- renderUI({
    req(data_shared())
    selectInput("vars_cfa", "Select Variables:", choices = names(data_shared()), multiple = TRUE)
  })
  
  output$varSelect_reg <- renderUI({
    req(data_shared())
    selectInput("vars_reg", "Select Variables:", choices = names(data_shared()), multiple = TRUE)
  })
  
  # Selected data for each tab
  selectedData_efa <- reactive({
    req(input$vars_efa)
    data_shared()[, input$vars_efa, drop = FALSE]
  })
  
  selectedData_cfa <- reactive({
    req(input$vars_cfa)
    data_shared()[, input$vars_cfa, drop = FALSE]
  })
  
  selectedData_reg <- reactive({
    req(input$vars_reg)
    data_shared()[, input$vars_reg, drop = FALSE]
  })
  
  # EFA Server Logic
  efa_result <- reactiveVal(NULL)
  parallel_result <- reactiveVal(NULL)
  
  observeEvent(input$analyze, {
    req(selectedData_efa())
    kmo <- KMO(selectedData_efa())
    bartlett <- cortest.bartlett(selectedData_efa())
    
    output$kmoBartlett <- renderPrint({
      cat("Kaiser-Meyer-Olkin (KMO) Test: ", round(kmo$MSA, 3), "\n")
      cat("Bartlett's Test of Sphericity: ", round(bartlett$p.value, 3), "\n")
    })
    
    # Correlation matrix with asterisks
    cor_test <- corr.test(selectedData_efa())
    r <- round(cor_test$r, 3)
    p <- cor_test$p
    
    stars <- matrix("", nrow = nrow(p), ncol = ncol(p))
    stars[p < 0.001] <- "***"
    stars[p < 0.01 & p >= 0.001] <- "**"
    stars[p < 0.05 & p >= 0.01] <- "*"
    
    r_formatted <- matrix("", nrow = nrow(r), ncol = ncol(r))
    for (i in 1:nrow(r)) {
      for (j in 1:ncol(r)) {
        r_formatted[i, j] <- paste0(formatC(r[i, j], format = "f", digits = 3), stars[i, j])
      }
    }
    
    rownames(r_formatted) <- rownames(r)
    colnames(r_formatted) <- colnames(r)
    
    output$corMatrix <- renderDT({
      datatable(as.data.frame(r_formatted), options = list(pageLength = 10, scrollX = TRUE))
    })
    
    # Scree Plot
    output$screePlot <- renderPlotly({
      ev <- eigen(cor(selectedData_efa(), use = "pairwise.complete.obs"))
      scree_data <- data.frame(Factor = 1:length(ev$values), Eigenvalue = ev$values)
      plot_ly(scree_data, x = ~Factor, y = ~Eigenvalue, type = 'scatter', mode = 'lines+markers',
              marker = list(size = 8)) %>%
        layout(title = "Scree Plot", xaxis = list(title = "Factor"), yaxis = list(title = "Eigenvalue"))
    })
    
    output$downloadScree <- downloadHandler(
      filename = function() { "scree_plot.png" },
      content = function(file) {
        ev <- eigen(cor(selectedData_efa(), use = "pairwise.complete.obs"))
        scree_data <- data.frame(Factor = 1:length(ev$values), Eigenvalue = ev$values)
        p <- ggplot(scree_data, aes(x = Factor, y = Eigenvalue)) +
          geom_point() + geom_line() +
          labs(title = "Scree Plot", x = "Factor", y = "Eigenvalue") +
          theme_minimal() + theme(plot.margin = margin(20, 20, 20, 20))
        ggsave(file, plot = p, width = 8, height = 6)
      }
    )
    
    # FA Parallel with interpretation
    output$faParallelPlot <- renderPlot({
      parallel <- fa.parallel(selectedData_efa(), fa = "fa", n.iter = 20, show.legend = TRUE)
      parallel_result(parallel)
    })
    
    output$parallelInterpretation <- renderPrint({
      req(parallel_result())
      parallel <- parallel_result()
      cat("Parallel Analysis Results:\n")
      cat("Suggested number of factors based on eigenvalues:", parallel$nfact, "\n")
      cat("Suggested number of components based on eigenvalues:", parallel$ncomp, "\n\n")
      cat("Interpretation:\n")
      cat("The parallel analysis suggests extracting", parallel$nfact, 
          "factors when the actual data eigenvalues are greater than the corresponding", 
          "percentiles of the random data eigenvalues.\n")
      cat("This is a more robust method than the traditional 'eigenvalue > 1' rule.\n")
    })
    
    output$downloadParallel <- downloadHandler(
      filename = function() { "fa_parallel_plot.png" },
      content = function(file) {
        png(file, width = 1000, height = 800)
        parallel <- fa.parallel(selectedData_efa(), fa = "fa", n.iter = 20, show.legend = TRUE)
        dev.off()
      }
    )
    
    efa <- fa(selectedData_efa(), nfactors = input$nfactors, rotate = input$rotate, fm = input$fm)
    efa_result(efa)
    
    output$efaResults <- renderPrint({
      print(efa, digits = 3)
    })
    
    output$varianceTable <- renderPrint({
      cat("Total Variance Explained:\n")
      print(round(efa$Vaccounted, 3))
    })
    
    output$reliabilityResult <- renderPrint({
      if (input$reliability == "Cronbach Alpha") {
        print(psych::alpha(selectedData_efa()), digits = 3)
      } else {
        print(omega(selectedData_efa(), nfactors = input$nfactors), digits = 3)
      }
    })
    
    output$downloadDiagram <- downloadHandler(
      filename = function() { "factor_diagram.png" },
      content = function(file) {
        png(file, width = 1000, height = 800)
        fa.diagram(efa_result())
        dev.off()
      }
    )
  })
  
  # Robust PDF download for EFA
  output$downloadAllEFA <- downloadHandler(
    filename = function() {
      paste0("EFA_Results_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".pdf")
    },
    content = function(file) {
      showNotification("Generating PDF report...", type = "message")
      
      tryCatch({
        # Create PDF
        pdf(file, width = 11, height = 8.5, paper = "a4")
        
        # Title page
        grid::grid.newpage()
        grid::pushViewport(grid::viewport(width = 0.9, height = 0.9))
        grid::grid.rect(gp = grid::gpar(fill = "white", col = NA))
        grid::grid.text("Exploratory Factor Analysis Report", 
                        y = 0.8, gp = grid::gpar(fontsize = 18, fontface = "bold"))
        grid::grid.text(paste("Generated on:", format(Sys.Date(), "%B %d, %Y")), 
                        y = 0.7, gp = grid::gpar(fontsize = 12))
        grid::grid.text("FAnalyzr - Statistical Analysis Tool", 
                        y = 0.6, gp = grid::gpar(fontsize = 10, col = "gray50"))
        
        # Analysis Parameters
        grid::grid.newpage()
        grid::grid.text("Analysis Parameters", 
                        y = 0.9, gp = grid::gpar(fontsize = 16, fontface = "bold"))
        
        params <- data.frame(
          Parameter = c("Extraction Method", "Rotation Method", "Number of Factors", "Reliability Method"),
          Value = c(input$fm, input$rotate, as.character(input$nfactors), input$reliability)
        )
        
        # Create table grob
        tg <- gridExtra::tableGrob(params, rows = NULL, 
                                   theme = gridExtra::ttheme_minimal(
                                     core = list(fg_params = list(fontsize = 10)),
                                     colhead = list(fg_params = list(fontsize = 11, fontface = "bold"))
                                   ))
        grid::grid.draw(tg)
        
        # KMO and Bartlett's Test
        if (!is.null(selectedData_efa())) {
          grid::grid.newpage()
          grid::grid.text("Data Suitability Tests", 
                          y = 0.9, gp = grid::gpar(fontsize = 16, fontface = "bold"))
          
          kmo <- KMO(selectedData_efa())
          bartlett <- cortest.bartlett(selectedData_efa())
          
          suitability_text <- c(
            paste("Kaiser-Meyer-Olkin (KMO) Test:", round(kmo$MSA, 3)),
            paste("Bartlett's Test of Sphericity p-value:", round(bartlett$p.value, 3)),
            "",
            "Interpretation:",
            ifelse(kmo$MSA > 0.8, "• KMO > 0.8: Marvelous", 
                   ifelse(kmo$MSA > 0.7, "• KMO > 0.7: Middling", 
                          ifelse(kmo$MSA > 0.6, "• KMO > 0.6: Mediocre", "• KMO < 0.6: Unacceptable"))),
            ifelse(bartlett$p.value < 0.05, "• Bartlett's test significant: Suitable for factor analysis", 
                   "• Bartlett's test not significant: Unsuitable for factor analysis")
          )
          
          for (i in 1:length(suitability_text)) {
            grid::grid.text(suitability_text[i], 
                            y = 0.8 - i * 0.04, 
                            x = 0.05, 
                            just = "left",
                            gp = grid::gpar(fontsize = 10))
          }
        }
        
        # Factor Loadings
        if (!is.null(efa_result())) {
          efa <- efa_result()
          
          grid::grid.newpage()
          grid::grid.text("Factor Loadings", 
                          y = 0.95, gp = grid::gpar(fontsize = 16, fontface = "bold"))
          
          # Convert loadings to data frame
          loadings_df <- as.data.frame(round(efa$loadings, 3))
          loadings_df <- cbind(Variable = rownames(loadings_df), loadings_df)
          rownames(loadings_df) <- NULL
          
          # Create table
          tg <- gridExtra::tableGrob(loadings_df, rows = NULL,
                                     theme = gridExtra::ttheme_minimal(
                                       core = list(fg_params = list(fontsize = 8)),
                                       colhead = list(fg_params = list(fontsize = 9, fontface = "bold"))
                                     ))
          grid::grid.draw(tg)
        }
        
        # Variance Explained
        if (!is.null(efa_result())) {
          grid::grid.newpage()
          grid::grid.text("Variance Explained", 
                          y = 0.95, gp = grid::gpar(fontsize = 16, fontface = "bold"))
          
          variance_df <- as.data.frame(round(efa_result()$Vaccounted, 3))
          variance_df <- cbind(Component = rownames(variance_df), variance_df)
          rownames(variance_df) <- NULL
          
          tg <- gridExtra::tableGrob(variance_df, rows = NULL,
                                     theme = gridExtra::ttheme_minimal(
                                       core = list(fg_params = list(fontsize = 9)),
                                       colhead = list(fg_params = list(fontsize = 10, fontface = "bold"))
                                     ))
          grid::grid.draw(tg)
        }
        
        # Reliability Results
        if (!is.null(selectedData_efa())) {
          grid::grid.newpage()
          grid::grid.text("Reliability Analysis", 
                          y = 0.9, gp = grid::gpar(fontsize = 16, fontface = "bold"))
          
          if (input$reliability == "Cronbach Alpha") {
            alpha_result <- psych::alpha(selectedData_efa())
            reliability_text <- c(
              paste("Cronbach's Alpha:", round(alpha_result$total$raw_alpha, 3)),
              paste("Number of Items:", ncol(selectedData_efa())),
              "",
              "Interpretation:",
              ifelse(alpha_result$total$raw_alpha > 0.9, "• Excellent reliability",
                     ifelse(alpha_result$total$raw_alpha > 0.8, "• Good reliability",
                            ifelse(alpha_result$total$raw_alpha > 0.7, "• Acceptable reliability",
                                   "• Poor reliability")))
            )
          } else {
            omega_result <- omega(selectedData_efa(), nfactors = input$nfactors)
            reliability_text <- c(
              paste("McDonald's Omega Total:", round(omega_result$omega.tot, 3)),
              paste("McDonald's Omega Hierarchical:", round(omega_result$omega_h, 3)),
              paste("Number of Items:", ncol(selectedData_efa()))
            )
          }
          
          for (i in 1:length(reliability_text)) {
            grid::grid.text(reliability_text[i], 
                            y = 0.8 - i * 0.04, 
                            x = 0.05, 
                            just = "left",
                            gp = grid::gpar(fontsize = 10))
          }
        }
        
        dev.off()
        showNotification("PDF report generated successfully!", type = "message")
        
      }, error = function(e) {
        showNotification(paste("Error generating PDF:", e$message), type = "error")
        if (exists("pdf", inherits = FALSE)) dev.off()
      })
    }
  )
  
  # CFA Server Logic
  modelResults <- eventReactive(input$runModel, {
    req(input$modelText, selectedData_cfa(), input$estimator)
    sem(model = input$modelText, data = selectedData_cfa(), estimator = input$estimator, std.lv = TRUE)
  })
  
  output$modelSummary <- renderPrint({
    req(modelResults())
    cat("Model Summary:\n")
    print(summary(modelResults(), fit.measures = TRUE))
    cat("\nFit Measures:\n")
    print(fitMeasures(modelResults()))
  })
  
  output$standardized <- renderPrint({
    req(modelResults())
    parameterEstimates(modelResults(), standardized = TRUE)[, c("lhs", "op", "rhs", "est", "std.all")]
  })
  
  output$unstandardized <- renderPrint({
    req(modelResults())
    parameterEstimates(modelResults(), standardized = FALSE)[, c("lhs", "op", "rhs", "est")]
  })
  
  output$modIndices <- renderDT({
    req(modelResults())
    mod_ind <- modindices(modelResults()) %>%
      mutate_if(is.numeric, ~ round(., 4))
    datatable(mod_ind, options = list(pageLength = 10, scrollX = TRUE))
  })
  
  output$reliabilityResults <- renderPrint({
    req(modelResults())
    
    reliability_vals <- tryCatch({
      semTools::reliability(modelResults())
    }, error = function(e) e)
    
    ave_vals <- tryCatch({
      semTools::AVE(modelResults())
    }, error = function(e) e)
    
    htmt_vals <- tryCatch({
      semTools::htmt(modelResults())
    }, error = function(e) e)
    
    omega_result <- tryCatch({
      psych::omega(selectedData_cfa(), warnings = FALSE)
    }, error = function(e) e)
    
    cat("Composite Reliability & Convergent Validity Measures:\n\n")
    
    cat("Composite Reliability (CR):\n")
    print(reliability_vals)
    
    cat("\nAverage Variance Extracted (AVE):\n")
    print(ave_vals)
    
    cat("\nHeterotrait-Monotrait Ratio (HTMT):\n")
    print(htmt_vals)
    
    cat("\nOmega Coefficients (Total and Hierarchical):\n")
    print(omega_result)
  })
  
  # Standardized Residual Covariances
  output$residCov <- renderPrint({
    req(modelResults())
    resid_cov <- resid(modelResults(), type = "standardized")$cov
    resid_cov <- round(resid_cov, 3)
    print(resid_cov)
  })
  
  # Discriminant Validity Calculation
  output$discriminantValidity <- renderTable({
    req(modelResults())
    
    # Get the standardized solution
    std_solution <- standardizedSolution(modelResults())
    
    # Extract factor correlations
    factor_cors <- std_solution %>%
      filter(op == "~~" & lhs != rhs & lhs %in% unique(std_solution$lhs[std_solution$op == "=~"])) %>%
      select(lhs, rhs, est.std) %>%
      rename(Factor1 = lhs, Factor2 = rhs, Correlation = est.std)
    
    # Calculate AVE for each factor
    loadings <- std_solution %>%
      filter(op == "=~") %>%
      group_by(lhs) %>%
      summarise(AVE = mean(est.std^2)) %>%
      rename(Factor = lhs)
    
    # Create discriminant validity table
    disc_validity <- factor_cors %>%
      left_join(loadings, by = c("Factor1" = "Factor")) %>%
      rename(AVE1 = AVE) %>%
      left_join(loadings, by = c("Factor2" = "Factor")) %>%
      rename(AVE2 = AVE) %>%
      mutate(Sqrt_AVE1 = sqrt(AVE1),
             Sqrt_AVE2 = sqrt(AVE2),
             Discriminant_Valid = ifelse(abs(Correlation) < pmin(Sqrt_AVE1, Sqrt_AVE2), "Yes", "No")) %>%
      select(Factor1, Factor2, Correlation, AVE1, AVE2, Discriminant_Valid)
    
    disc_validity <- disc_validity %>%
      mutate(across(where(is.numeric), ~ round(., 3)))
    
    disc_validity
  }, rownames = FALSE)
  
  # Robust PDF download for CFA
  output$downloadAllCFA <- downloadHandler(
    filename = function() {
      paste0("CFA_Results_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".pdf")
    },
    content = function(file) {
      showNotification("Generating PDF report...", type = "message")
      
      tryCatch({
        pdf(file, width = 11, height = 8.5, paper = "a4")
        
        # Title page
        grid::grid.newpage()
        grid::pushViewport(grid::viewport(width = 0.9, height = 0.9))
        grid::grid.rect(gp = grid::gpar(fill = "white", col = NA))
        grid::grid.text("Confirmatory Factor Analysis Report", 
                        y = 0.8, gp = grid::gpar(fontsize = 18, fontface = "bold"))
        grid::grid.text(paste("Generated on:", format(Sys.Date(), "%B %d, %Y")), 
                        y = 0.7, gp = grid::gpar(fontsize = 12))
        grid::grid.text("FAnalyzr - Statistical Analysis Tool", 
                        y = 0.6, gp = grid::gpar(fontsize = 10, col = "gray50"))
        
        # Model Specification
        grid::grid.newpage()
        grid::grid.text("Model Specification", 
                        y = 0.95, gp = grid::gpar(fontsize = 16, fontface = "bold"))
        
        if (!is.null(input$modelText) && nchar(input$modelText) > 0) {
          model_lines <- strsplit(input$modelText, "\n")[[1]]
          for (i in 1:length(model_lines)) {
            grid::grid.text(model_lines[i], 
                            y = 0.9 - i * 0.03, 
                            x = 0.05, 
                            just = "left",
                            gp = grid::gpar(fontsize = 9, family = "mono"))
          }
        }
        
        # Fit Measures
        if (!is.null(modelResults())) {
          grid::grid.newpage()
          grid::grid.text("Model Fit Measures", 
                          y = 0.95, gp = grid::gpar(fontsize = 16, fontface = "bold"))
          
          fit <- fitMeasures(modelResults())
          fit_measures <- c("chisq", "df", "pvalue", "cfi", "tli", "rmsea", "srmr")
          fit_names <- c("Chi-square", "Degrees of Freedom", "P-value", "CFI", "TLI", "RMSEA", "SRMR")
          
          fit_df <- data.frame(
            Measure = fit_names,
            Value = round(fit[fit_measures], 3)
          )
          
          tg <- gridExtra::tableGrob(fit_df, rows = NULL,
                                     theme = gridExtra::ttheme_minimal(
                                       core = list(fg_params = list(fontsize = 10)),
                                       colhead = list(fg_params = list(fontsize = 11, fontface = "bold"))
                                     ))
          grid::grid.draw(tg)
          
          # Fit interpretation
          grid::grid.text("Fit Interpretation:", 
                          y = 0.4, x = 0.05, just = "left",
                          gp = grid::gpar(fontsize = 10, fontface = "bold"))
          
          interpretation <- character()
          if (fit["cfi"] > 0.95 && fit["rmsea"] < 0.06) {
            interpretation <- c(interpretation, "• Excellent model fit (CFI > 0.95, RMSEA < 0.06)")
          } else if (fit["cfi"] > 0.90 && fit["rmsea"] < 0.08) {
            interpretation <- c(interpretation, "• Acceptable model fit (CFI > 0.90, RMSEA < 0.08)")
          } else {
            interpretation <- c(interpretation, "• Poor model fit - consider revising your model")
          }
          
          if (fit["pvalue"] < 0.05) {
            interpretation <- c(interpretation, "• Significant chi-square test (p < 0.05)")
          } else {
            interpretation <- c(interpretation, "• Non-significant chi-square test (p >= 0.05)")
          }
          
          for (i in 1:length(interpretation)) {
            grid::grid.text(interpretation[i], 
                            y = 0.35 - i * 0.03, 
                            x = 0.05, 
                            just = "left",
                            gp = grid::gpar(fontsize = 9))
          }
        }
        
        # Standardized Estimates
        if (!is.null(modelResults())) {
          grid::grid.newpage()
          grid::grid.text("Standardized Estimates", 
                          y = 0.95, gp = grid::gpar(fontsize = 16, fontface = "bold"))
          
          std_est <- parameterEstimates(modelResults(), standardized = TRUE)
          std_est <- std_est[, c("lhs", "op", "rhs", "est", "std.all", "pvalue")]
          std_est <- std_est[std_est$op %in% c("=~", "~~", "~"), ]
          std_est <- std_est[!duplicated(std_est), ]
          
          if (nrow(std_est) > 0) {
            std_est$est <- round(std_est$est, 3)
            std_est$std.all <- round(std_est$std.all, 3)
            std_est$pvalue <- round(std_est$pvalue, 3)
            
            # Display in chunks if too many rows
            if (nrow(std_est) <= 20) {
              tg <- gridExtra::tableGrob(std_est, rows = NULL,
                                         theme = gridExtra::ttheme_minimal(
                                           core = list(fg_params = list(fontsize = 8)),
                                           colhead = list(fg_params = list(fontsize = 9, fontface = "bold"))
                                         ))
              grid::grid.draw(tg)
            } else {
              # Show first 20 rows with note
              tg <- gridExtra::tableGrob(std_est[1:20, ], rows = NULL,
                                         theme = gridExtra::ttheme_minimal(
                                           core = list(fg_params = list(fontsize = 8)),
                                           colhead = list(fg_params = list(fontsize = 9, fontface = "bold"))
                                         ))
              grid::grid.draw(tg)
              grid::grid.text("Note: Showing first 20 of many estimates", 
                              y = 0.1, gp = grid::gpar(fontsize = 9, col = "gray50"))
            }
          }
        }
        
        dev.off()
        showNotification("PDF report generated successfully!", type = "message")
        
      }, error = function(e) {
        showNotification(paste("Error generating PDF:", e$message), type = "error")
        if (exists("pdf", inherits = FALSE)) dev.off()
      })
    }
  )
  
  output$semPlot <- renderPlot({
    req(modelResults())
    semPaths(modelResults(), what = "std", layout = input$diagram_layout_cfa, edge.label.cex = 1.2, 
             sizeMan = input$node_size_cfa, sizeLat = input$node_size_cfa, nCharNodes = 0, edge.color = "black", 
             mar = c(5, 5, 5, 5), fade = FALSE, rotation = 2, label.cex = 1.2, 
             edge.width = 2, style = "ram", nodeWidth = 3, edge.arrow.size = 0.5)
  })
  
  output$downloadPlot <- downloadHandler(
    filename = function() {
      paste("sem_plot", ".png", sep = "")
    },
    content = function(file) {
      png(file, width = 1200, height = 900)
      semPaths(modelResults(), what = "std", layout = input$diagram_layout_cfa, edge.label.cex = 1.2, 
               sizeMan = input$node_size_cfa, sizeLat = input$node_size_cfa, nCharNodes = 0, edge.color = "black", 
               mar = c(5, 5, 5, 5), fade = FALSE, rotation = 2, label.cex = 1.2, 
               edge.width = 2, style = "ram", nodeWidth = 3, edge.arrow.size = 0.5)
      dev.off()
    }
  )
  
  # Regression Table Server Logic
  regResults <- eventReactive(input$runRegModel, {
    req(input$regModelText, selectedData_reg(), input$regEstimator)
    sem(model = input$regModelText, data = selectedData_reg(), estimator = input$regEstimator)
  })
  
  output$regressionTable <- renderDT({
    req(regResults())
    params <- parameterEstimates(regResults(), standardized = TRUE)
    effects_table <- params[params$op == "~", c("lhs", "op", "rhs", "est", "se", "z", "pvalue")]
    colnames(effects_table) <- c("Dependent", "Operator", "Predictor", "Estimate", "Std. Error", "Z-value", "P-value")
    
    effects_table$Estimate <- round(effects_table$Estimate, 4)
    effects_table$`Std. Error` <- round(effects_table$`Std. Error`, 4)
    effects_table$`Z-value` <- round(effects_table$`Z-value`, 4)
    effects_table$`P-value` <- round(effects_table$`P-value`, 4)
    
    datatable(effects_table, options = list(pageLength = 15, scrollX = TRUE))
  })
  
  # Robust PDF download for Regression
  output$downloadAllReg <- downloadHandler(
    filename = function() {
      paste0("Regression_Results_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".pdf")
    },
    content = function(file) {
      showNotification("Generating PDF report...", type = "message")
      
      tryCatch({
        pdf(file, width = 11, height = 8.5, paper = "a4")
        
        # Title page
        grid::grid.newpage()
        grid::pushViewport(grid::viewport(width = 0.9, height = 0.9))
        grid::grid.rect(gp = grid::gpar(fill = "white", col = NA))
        grid::grid.text("Regression Analysis Report", 
                        y = 0.8, gp = grid::gpar(fontsize = 18, fontface = "bold"))
        grid::grid.text(paste("Generated on:", format(Sys.Date(), "%B %d, %Y")), 
                        y = 0.7, gp = grid::gpar(fontsize = 12))
        grid::grid.text("FAnalyzr - Statistical Analysis Tool", 
                        y = 0.6, gp = grid::gpar(fontsize = 10, col = "gray50"))
        
        # Model Specification
        grid::grid.newpage()
        grid::grid.text("Regression Model Specification", 
                        y = 0.95, gp = grid::gpar(fontsize = 16, fontface = "bold"))
        
        if (!is.null(input$regModelText) && nchar(input$regModelText) > 0) {
          model_lines <- strsplit(input$regModelText, "\n")[[1]]
          for (i in 1:length(model_lines)) {
            grid::grid.text(model_lines[i], 
                            y = 0.9 - i * 0.03, 
                            x = 0.05, 
                            just = "left",
                            gp = grid::gpar(fontsize = 9, family = "mono"))
          }
        }
        
        # Regression Results
        if (!is.null(regResults())) {
          grid::grid.newpage()
          grid::grid.text("Regression Coefficients", 
                          y = 0.95, gp = grid::gpar(fontsize = 16, fontface = "bold"))
          
          params <- parameterEstimates(regResults(), standardized = TRUE)
          effects_table <- params[params$op == "~", c("lhs", "op", "rhs", "est", "se", "z", "pvalue", "std.all")]
          colnames(effects_table) <- c("Dependent", "Operator", "Predictor", "Estimate", "Std. Error", "Z-value", "P-value", "Std. Estimate")
          
          effects_table <- effects_table %>%
            mutate(across(where(is.numeric), ~ round(., 3)))
          
          if (nrow(effects_table) > 0) {
            tg <- gridExtra::tableGrob(effects_table, rows = NULL,
                                       theme = gridExtra::ttheme_minimal(
                                         core = list(fg_params = list(fontsize = 8)),
                                         colhead = list(fg_params = list(fontsize = 9, fontface = "bold"))
                                       ))
            grid::grid.draw(tg)
          }
        }
        
        # Model Fit (if available)
        if (!is.null(regResults()) && !is.null(fitMeasures(regResults()))) {
          grid::grid.newpage()
          grid::grid.text("Model Fit Measures", 
                          y = 0.95, gp = grid::gpar(fontsize = 16, fontface = "bold"))
          
          fit <- fitMeasures(regResults())
          fit_measures <- c("chisq", "df", "pvalue", "cfi", "tli", "rmsea", "srmr")
          fit_names <- c("Chi-square", "Degrees of Freedom", "P-value", "CFI", "TLI", "RMSEA", "SRMR")
          
          fit_df <- data.frame(
            Measure = fit_names,
            Value = round(fit[fit_measures], 3)
          )
          
          tg <- gridExtra::tableGrob(fit_df, rows = NULL,
                                     theme = gridExtra::ttheme_minimal(
                                       core = list(fg_params = list(fontsize = 10)),
                                       colhead = list(fg_params = list(fontsize = 11, fontface = "bold"))
                                     ))
          grid::grid.draw(tg)
        }
        
        dev.off()
        showNotification("PDF report generated successfully!", type = "message")
        
      }, error = function(e) {
        showNotification(paste("Error generating PDF:", e$message), type = "error")
        if (exists("pdf", inherits = FALSE)) dev.off()
      })
    }
  )
}

shinyApp(ui, server)
