# Load required libraries
library(shiny)
library(ggplot2)
library(pROC)
library(MLmetrics)
library(DT)
library(caret)
library(dplyr)
library(readxl)
library(haven)
library(shinythemes)
library(officer)
library(shinycssloaders)
library(shinyjs)

ui <- fluidPage(
  theme = shinytheme("flatly"),
  useShinyjs(),
  
  tags$head(
    tags$style(HTML("
      /* Main styling */
      body {
        background-color: #f8fafc;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        line-height: 1.6;
        margin: 0;
        padding: 0;
        display: flex;
        flex-direction: column;
        min-height: 100vh;
      }
      
      .container-fluid {
        flex: 1;
      }
      
      /* Header styling */
      .header {
        background: linear-gradient(135deg, #3498db, #2c3e50);
        color: white;
        padding: 25px 0;
        margin-bottom: 0;
        border-radius: 0;
        box-shadow: 0 4px 20px rgba(0,0,0,0.15);
        position: relative;
        overflow: hidden;
      }
      
      .header::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background: url('data:image/svg+xml,<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 100 100\" preserveAspectRatio=\"none\"><path d=\"M0,0 L100,0 L100,100 Z\" fill=\"rgba(255,255,255,0.05)\"/></svg>');
        background-size: cover;
      }
      
      .app-title {
        font-size: 32px;
        font-weight: 700;
        margin: 0;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        text-align: center;
        position: relative;
      }
      
      .app-subtitle {
        font-size: 16px;
        opacity: 0.95;
        margin-top: 8px;
        text-align: center;
        font-weight: 300;
        position: relative;
      }
      
      /* Navigation styling */
      .nav-tabs {
        background: white;
        border-bottom: 3px solid #3498db;
        box-shadow: 0 2px 10px rgba(0,0,0,0.08);
      }
      
      .nav-tabs > li > a {
        color: #2c3e50;
        font-weight: 600;
        border: none !important;
        border-radius: 0;
        padding: 15px 30px;
        transition: all 0.3s ease;
        position: relative;
        overflow: hidden;
      }
      
      .nav-tabs > li > a::before {
        content: '';
        position: absolute;
        bottom: 0;
        left: 50%;
        width: 0;
        height: 3px;
        background: #3498db;
        transition: all 0.3s ease;
        transform: translateX(-50%);
      }
      
      .nav-tabs > li > a:hover::before {
        width: 80%;
      }
      
      .nav-tabs > li.active > a {
        color: #3498db !important;
        background: transparent !important;
      }
      
      .nav-tabs > li.active > a::before {
        width: 100%;
      }
      
      /* Panel styling */
      .sidebar-panel {
        background-color: white;
        border-radius: 12px;
        padding: 25px !important;
        box-shadow: 0 4px 15px rgba(0,0,0,0.08);
        border: 1px solid #e8ecef;
        transition: transform 0.3s ease, box-shadow 0.3s ease;
      }
      
      .sidebar-panel:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(0,0,0,0.12);
      }
      
      .main-panel {
        background-color: white;
        border-radius: 12px;
        padding: 30px !important;
        box-shadow: 0 4px 15px rgba(0,0,0,0.08);
        border: 1px solid #e8ecef;
        min-height: 600px;
        animation: fadeIn 0.8s ease-in;
      }
      
      @keyframes fadeIn {
        from { opacity: 0; transform: translateY(20px); }
        to { opacity: 1; transform: translateY(0); }
      }
      
      /* Button styling */
      .btn-primary {
        background: linear-gradient(135deg, #3498db, #2980b9);
        border: none;
        font-weight: 600;
        border-radius: 8px;
        padding: 12px 20px;
        transition: all 0.3s ease;
        box-shadow: 0 2px 8px rgba(52, 152, 219, 0.3);
      }
      
      .btn-primary:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 15px rgba(52, 152, 219, 0.4);
        background: linear-gradient(135deg, #2980b9, #2573a7);
      }
      
      .btn-success {
        background: linear-gradient(135deg, #2ecc71, #27ae60);
        border: none;
        font-weight: 600;
        border-radius: 8px;
        padding: 12px 20px;
        transition: all 0.3s ease;
        box-shadow: 0 2px 8px rgba(46, 204, 113, 0.3);
      }
      
      .btn-success:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 15px rgba(46, 204, 113, 0.4);
        background: linear-gradient(135deg, #27ae60, #229954);
      }
      
      .btn-info {
        background: linear-gradient(135deg, #17a2b8, #138496);
        border: none;
        font-weight: 600;
        border-radius: 8px;
        padding: 12px 20px;
        transition: all 0.3s ease;
        box-shadow: 0 2px 8px rgba(23, 162, 184, 0.3);
      }
      
      .btn-info:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 15px rgba(23, 162, 184, 0.4);
        background: linear-gradient(135deg, #138496, #117a8b);
      }
      
      /* Input styling */
      .shiny-input-container {
        margin-bottom: 20px;
      }
      
      .form-control {
        border-radius: 8px;
        border: 2px solid #e8ecef;
        padding: 10px 15px;
        transition: all 0.3s ease;
        font-size: 14px;
      }
      
      .form-control:focus {
        border-color: #3498db;
        box-shadow: 0 0 0 3px rgba(52, 152, 219, 0.1);
      }
      
      .selectize-input {
        border-radius: 8px !important;
        border: 2px solid #e8ecef !important;
        padding: 10px 15px !important;
        box-shadow: none !important;
        transition: all 0.3s ease !important;
      }
      
      .selectize-input.focus {
        border-color: #3498db !important;
        box-shadow: 0 0 0 3px rgba(52, 152, 219, 0.1) !important;
      }
      
      /* Text styling */
      h2, h3, h4 {
        color: #2c3e50;
        font-weight: 600;
        margin-bottom: 20px;
      }
      
      h4 {
        margin-top: 25px;
        border-bottom: 2px solid #f0f3f5;
        padding-bottom: 10px;
        position: relative;
      }
      
      h4::after {
        content: '';
        position: absolute;
        bottom: -2px;
        left: 0;
        width: 50px;
        height: 2px;
        background: #3498db;
      }
      
      .input-info, .notice {
        font-size: 0.9em;
        color: #6c757d;
        background: #f8f9fa;
        padding: 10px 15px;
        border-radius: 8px;
        border-left: 4px solid #3498db;
      }
      
      .instruction {
        font-size: 0.85em;
        color: #6c757d;
        margin-bottom: 15px;
        padding: 8px 12px;
        background: #fff3cd;
        border-radius: 6px;
        border-left: 4px solid #ffc107;
      }
      
      .warning-text {
        color: #dc3545;
        font-weight: 600;
        background-color: #f8d7da;
        padding: 10px 15px;
        border-radius: 8px;
        display: block;
        margin: 10px 0;
        border-left: 4px solid #dc3545;
        animation: pulse 2s infinite;
      }
      
      .success-text {
        color: #155724;
        font-weight: 600;
        background-color: #d4edda;
        padding: 10px 15px;
        border-radius: 8px;
        display: block;
        margin: 10px 0;
        border-left: 4px solid #28a745;
      }
      
      @keyframes pulse {
        0% { transform: scale(1); }
        50% { transform: scale(1.02); }
        100% { transform: scale(1); }
      }
      
      /* Footer styling */
      .footer {
        color: #6c757d;
        font-style: italic;
        text-align: center;
        padding-top: 25px;
        margin-top: 30px;
        border-top: 1px solid #e8ecef;
        font-size: 0.9em;
      }
      
      /* Copyright footer styling */
      .copyright-footer {
        background: linear-gradient(135deg, #2c3e50, #34495e);
        color: white;
        padding: 20px 0;
        margin-top: 40px;
        text-align: center;
        font-size: 0.85em;
        border-top: 3px solid #3498db;
        width: 100%;
      }
      
      .copyright-content {
        max-width: 1200px;
        margin: 0 auto;
        padding: 0 20px;
      }
      
      .copyright-text {
        margin: 5px 0;
        opacity: 0.9;
      }
      
      .copyright-notice {
        font-weight: 600;
        color: #3498db;
        margin-bottom: 8px;
      }
      
      .copyright-links {
        margin-top: 10px;
      }
      
      .copyright-links a {
        color: #ecf0f1;
        margin: 0 10px;
        text-decoration: none;
        transition: color 0.3s ease;
      }
      
      .copyright-links a:hover {
        color: #3498db;
        text-decoration: underline;
      }
      
      /* Table styling */
      .dataTables_wrapper {
        margin-top: 20px;
      }
      
      /* Home page specific */
      .feature-card {
        background: white;
        border-radius: 12px;
        padding: 25px;
        margin-bottom: 20px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.08);
        border: 1px solid #e8ecef;
        transition: all 0.3s ease;
        text-align: center;
      }
      
      .feature-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 8px 25px rgba(0,0,0,0.15);
      }
      
      .feature-icon {
        font-size: 48px;
        color: #3498db;
        margin-bottom: 15px;
      }
      
      /* About page specific */
      .about-card {
        background: white;
        border-radius: 12px;
        padding: 30px;
        margin-bottom: 25px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.08);
        border: 1px solid #e8ecef;
      }
      
      .developer-info {
        text-align: center;
        padding: 20px;
      }
      
      .developer-avatar {
        width: 120px;
        height: 120px;
        border-radius: 50%;
        margin: 0 auto 20px;
        background: linear-gradient(135deg, #3498db, #2c3e50);
        display: flex;
        align-items: center;
        justify-content: center;
        color: white;
        font-size: 48px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.2);
      }
      
      /* Loading animations */
      .loading-spinner {
        color: #3498db;
        font-size: 18px;
        text-align: center;
        padding: 20px;
      }
      
      /* Responsive adjustments */
      @media (max-width: 768px) {
        .sidebar-panel, .main-panel {
          padding: 20px !important;
        }
        
        .app-title {
          font-size: 24px;
        }
        
        .nav-tabs > li > a {
          padding: 12px 20px;
        }
        
        .copyright-footer {
          padding: 15px 0;
          font-size: 0.8em;
        }
      }
    "))
  ),
  
  # Header Section
  div(class = "header",
      div(class = "container-fluid",
          div(
            h1(class = "app-title", "ROC Curve Builder"),
            p(class = "app-subtitle", "Professional Binary Classification Performance Analysis Tool")
          )
      )
  ),
  
  # Navigation Tabs
  navbarPage(
    title = "",
    id = "main_nav",
    theme = shinytheme("flatly"),
    footer = div(
      class = "copyright-footer",
      div(class = "copyright-content",
          div(class = "copyright-notice", 
              icon("copyright"), 
              "Copyright 2025 Mudasir Mohammed Ibrahim. All Rights Reserved."),
          div(class = "copyright-links",
              a("Terms of Use", href = "#", onclick = "alert('Terms of Use: This software is provided for academic and research purposes.')"),
              "|",
              a("Privacy Policy", href = "#", onclick = "alert('Privacy Policy: This application does not collect or store any user data. All processing occurs locally in your browser.')"),
              "|",
              a("Contact", href = "mailto:mudassiribrahim30@gmail.com"),
              "|",
              a("License Information", href = "#", onclick = "alert('License: Academic and Research Use.')")
          )
      )
    ),
    
    # Home Tab
    tabPanel(
      "Home",
      icon = icon("home"),
      div(class = "container-fluid",
          style = "padding: 30px;",
          div(class = "main-panel",
              h2("Welcome to ROC Curve Builder", style = "text-align: center; color: #2c3e50;"),
              p("A comprehensive tool for building, analyzing, and interpreting Receiver Operating Characteristic (ROC) curves for binary classification models.", 
                style = "text-align: center; font-size: 16px; color: #6c757d; margin-bottom: 40px;"),
              
              fluidRow(
                column(4,
                       div(class = "feature-card",
                           div(class = "feature-icon", icon("chart-line")),
                           h4("ROC Curve Visualization"),
                           p("Create ROC curves for presentations and publications.")
                       )
                ),
                column(4,
                       div(class = "feature-card",
                           div(class = "feature-icon", icon("calculator")),
                           h4("Performance Metrics"),
                           p("Calculate comprehensive performance metrics including AUC, accuracy, sensitivity, specificity, and confusion matrices.")
                       )
                ),
                column(4,
                       div(class = "feature-card",
                           div(class = "feature-icon", icon("download")),
                           h4("Export Capabilities"),
                           p("Download high-quality plots and comprehensive reports in multiple formats for documentation and sharing.")
                       )
                )
              ),
              
              fluidRow(
                column(6,
                       div(class = "feature-card",
                           div(class = "feature-icon", icon("database")),
                           h4("Multiple Data Formats"),
                           p("Support for various data formats including CSV, Excel, Stata, and SPSS files with automatic data validation.")
                       )
                ),
                column(6,
                       div(class = "feature-card",
                           div(class = "feature-icon", icon("sliders-h")),
                           h4("Customizable Analysis"),
                           p("Flexible variable selection, recoding options, and graph customization to suit your specific analysis needs.")
                       )
                )
              ),
              
              div(style = "text-align: center; margin-top: 40px;",
                  actionButton("goto_analysis", "Start Analysis", 
                               class = "btn-primary btn-lg",
                               icon = icon("rocket"),
                               style = "padding: 15px 30px; font-size: 16px;")
              ),
              
              h3("How to Use", style = "margin-top: 50px;"),
              tags$ol(
                style = "font-size: 15px; color: #495057; line-height: 1.8;",
                tags$li("Navigate to the 'Build ROC Curve' tab"),
                tags$li("Upload your dataset in supported format (CSV, Excel, Stata, SPSS)"),
                tags$li("Select your binary outcome variable and predictor variables"),
                tags$li("Recode variables if necessary (ensure outcome is 0/1)"),
                tags$li("Click 'Generate ROC Curve' to perform analysis"),
                tags$li("Customize graph appearance and download results")
              )
          )
      )
    ),
    
    # Build ROC Curve Tab
    tabPanel(
      "Build ROC Curve",
      icon = icon("chart-line"),
      div(class = "container-fluid",
          sidebarLayout(
            sidebarPanel(
              class = "sidebar-panel",
              width = 4,
              h3("Data Input", style = "margin-top: 0;"),
              
              fileInput("data", "Upload Your File",
                        accept = c(".csv", ".xlsx", ".xls", ".sav", ".dta"),
                        buttonLabel = "Browse Files...",
                        placeholder = "No file selected",
                        width = "100%"),
              div(class = "input-info", "Supported formats: CSV, Excel (XLS/XLSX), Stata (DTA), SPSS (SAV)"),
              
              # Data upload status
              uiOutput("upload_status"),
              
              div(class = "notice", 
                  icon("info-circle"), 
                  strong(" Note:"), 
                  " After uploading, click 'Generate ROC Curve' to populate variable selectors."),
              br(),
              
              h4("Variable Selection"),
              selectInput("outcome", "Outcome Variable", choices = NULL, width = "100%"),
              uiOutput("outcome_warning"),
              selectizeInput("predictors", "Predictor Variable(s)", choices = NULL, multiple = TRUE, width = "100%"),
              
              div(class = "instruction", 
                  icon("exclamation-triangle"), 
                  "Ensure outcome is binary (0/1). Recode if needed:"),
              uiOutput("recode_ui"),
              
              h4("Graph Customization"),
              textInput("graph_title", "Graph Title", "ROC Curve", width = "100%"),
              textInput("x_label", "X-axis Label", "False Positive Rate", width = "100%"),
              textInput("y_label", "Y-axis Label", "True Positive Rate", width = "100%"),
              
              # Advanced customization options
              h4("Advanced Customization"),
              selectInput("color_palette", "Color Palette", 
                          choices = c("Blue Gradient" = "blue", 
                                      "Red Gradient" = "red",
                                      "Green Gradient" = "green",
                                      "Purple Gradient" = "purple",
                                      "Custom" = "custom"),
                          selected = "blue"),
              conditionalPanel(
                condition = "input.color_palette == 'custom'",
                textInput("custom_color", "Custom Color", value = "#3498db")
              ),
              sliderInput("line_size", "Line Size", min = 0.5, max = 3, value = 1.5, step = 0.1),
              sliderInput("alpha_level", "Transparency", min = 0.1, max = 1, value = 0.7, step = 0.1),
              checkboxInput("show_auc", "Show AUC on Plot", value = TRUE),
              checkboxInput("show_diagonal", "Show Diagonal Reference", value = TRUE),
              
              h4("Actions"),
              div(style = "text-align: center;",
                  actionButton("generate", "Generate ROC Curve", 
                               class = "btn-primary",
                               icon = icon("play"),
                               style = "width: 100%; margin-bottom: 10px;"),
                  downloadButton("download_plot", "Download Plot (PNG)", 
                                 class = "btn-success",
                                 style = "width: 100%; margin-bottom: 10px;"),
                  downloadButton("download_word", "Download Report (DOCX)", 
                                 class = "btn-info",
                                 style = "width: 100%;")
              )
            ),
            
            mainPanel(
              class = "main-panel",
              width = 8,
              withSpinner(
                plotOutput("roc_plot", height = "500px"),
                type = 4, color = "#3498db", size = 1
              ),
              
              h4("Performance Metrics"),
              DTOutput("metrics_table"),
              
              h4("Confusion Matrix"),
              DTOutput("conf_matrix_table"),
              
              div(class = "footer",
                  "ROC Curve Builder v2.0 | ",
                  icon("envelope"), " ",
                  a("mudassiribrahim30@gmail.com", href = "mailto:mudassiribrahim30@gmail.com"),
                  " | ",
                  icon("github"), " ",
                  a("View on GitHub", href = "https://github.com", target = "_blank")
              )
            )
          )
      )
    ),
    
    # About Tab
    tabPanel(
      "About",
      icon = icon("info-circle"),
      div(class = "container-fluid",
          style = "padding: 30px;",
          div(class = "main-panel",
              h2("About ROC Curve Builder", style = "text-align: center;"),
              
              fluidRow(
                column(8, offset = 2,
                       div(class = "about-card",
                           h3("Application Overview"),
                           p("The ROC Curve Builder is a professional-grade Shiny application designed for researchers, data scientists, and analysts working with binary classification models."),
                           
                           h4("Key Features"),
                           tags$ul(
                             tags$li("Interactive ROC curve visualization with smooth animations"),
                             tags$li("Comprehensive performance metrics calculation"),
                             tags$li("Support for multiple data formats"),
                             tags$li("Professional reporting and export capabilities"),
                             tags$li("User-friendly interface with real-time validation")
                           ),
                           
                           h4("Technical Specifications"),
                           p("Built with R Shiny framework, utilizing state-of-the-art packages for statistical analysis and visualization."),
                           
                           h4("Intended Use"),
                           p("This tool is designed for academic research, clinical studies, machine learning model evaluation, and any scenario requiring binary classification performance assessment.")
                       )
                )
              ),
              
              fluidRow(
                column(6, offset = 3,
                       div(class = "about-card developer-info",
                           h3("Developer Information"),
                           div(class = "developer-avatar",
                               icon("user")
                           ),
                           h4("Mudasir Mohammed Ibrahim"),
                           p("Registered Nurse"),
                           tags$p(icon("envelope"), " ", 
                                  a("mudassiribrahim30@gmail.com", href = "mailto:mudassiribrahim30@gmail.com")),
                           tags$p(icon("github"), " ", 
                                  a("GitHub Profile", href = "https://github.com", target = "_blank")),
                           tags$p(icon("linkedin"), " ", 
                                  a("LinkedIn Profile", href = "https://linkedin.com", target = "_blank")),
                           
                           div(style = "margin-top: 20px; padding: 15px; background: #f8f9fa; border-radius: 8px;",
                               h5("Copyright Notice"),
                               p(icon("copyright"), "Copyright 2025 Mudasir Mohammed Ibrahim. All Rights Reserved.", 
                                 style = "font-size: 0.9em; color: #6c757d; margin-bottom: 5px;"),
                               p("",
                                 style = "font-size: 0.8em; color: #6c757d;")
                           )
                       )
                )
              ),
              
              div(class = "about-card",
                  h3("Citation"),
                  p("If you use this application in your research or publications, please cite:"),
                  p(em("Ibrahim, M. M. (2025). ROC Curve Builder: A Shiny Application for Binary Classification Performance Analysis. [Computer software]")),
                  
                  h3("License"),
                  p("This software is provided for academic and research purposes."),
                  
                  h3("Version"),
                  p("Current Version: 2.0 | Release Date: November, 2025")
              )
          )
      )
    )
  )
)

server <- function(input, output, session) {
  
  # Reactive value to store data
  data_reactive <- reactiveVal(NULL)
  upload_status <- reactiveVal("No file uploaded")
  
  # Enhanced data reading function with better error handling
  observeEvent(input$data, {
    req(input$data)
    
    tryCatch({
      # Show loading message
      upload_status("Reading file...")
      
      ext <- tools::file_ext(input$data$datapath)
      
      df <- switch(tolower(ext),
                   csv = {
                     # Enhanced CSV reading with multiple options
                     result <- tryCatch({
                       read.csv(input$data$datapath, stringsAsFactors = FALSE, fileEncoding = "UTF-8")
                     }, error = function(e) {
                       # Try with different encoding
                       read.csv(input$data$datapath, stringsAsFactors = FALSE, fileEncoding = "latin1")
                     })
                     result
                   },
                   xlsx = read_excel(input$data$datapath),
                   xls = read_excel(input$data$datapath),
                   sav = read_sav(input$data$datapath),
                   dta = read_dta(input$data$datapath),
                   stop("Unsupported file type: ", ext)
      )
      
      # Validate that data was read successfully
      if (is.null(df)) {
        stop("Failed to read data from file")
      }
      
      if (nrow(df) == 0) {
        stop("The uploaded file contains no data")
      }
      
      if (ncol(df) == 0) {
        stop("The uploaded file contains no columns")
      }
      
      # Store the data
      data_reactive(df)
      upload_status(paste("Successfully loaded", nrow(df), "rows and", ncol(df), "columns"))
      
      # Update variable selectors
      updateSelectInput(session, "outcome", choices = names(df))
      updateSelectizeInput(session, "predictors", choices = names(df), server = TRUE)
      
    }, error = function(e) {
      # Set error status
      error_msg <- paste("Error reading file:", e$message)
      upload_status(error_msg)
      data_reactive(NULL)
      
      # Reset variable selectors
      updateSelectInput(session, "outcome", choices = NULL)
      updateSelectizeInput(session, "predictors", choices = NULL, server = TRUE)
      
      # Show error notification
      showNotification(error_msg, type = "error", duration = 10)
    })
  })
  
  # Display upload status
  output$upload_status <- renderUI({
    status <- upload_status()
    if (grepl("Error", status)) {
      div(class = "warning-text", icon("exclamation-triangle"), status)
    } else if (grepl("Successfully", status)) {
      div(class = "success-text", icon("check-circle"), status)
    } else {
      div(class = "notice", icon("info-circle"), status)
    }
  })
  
  output$outcome_warning <- renderUI({
    req(input$outcome, input$data)
    df <- data_reactive()
    if (is.null(df) || !input$outcome %in% names(df)) return(NULL)
    
    vals <- unique(na.omit(as.character(df[[input$outcome]])))
    if (length(vals) > 2) {
      HTML("<div class='warning-text'><i class='fa fa-exclamation-triangle'></i> Warning: Outcome has >2 levels. Please select a binary variable or use recoding.</div>")
    } else {
      NULL
    }
  })
  
  output$recode_ui <- renderUI({
    req(input$outcome, input$data)
    df <- data_reactive()
    if (is.null(df) || !input$outcome %in% names(df)) return(NULL)
    
    choices <- unique(as.character(na.omit(df[[input$outcome]])))
    if (length(choices) == 0) return(NULL)
    
    tagList(
      selectInput("old_val1", "Original Value 1", choices = choices, selected = choices[1], width = "100%"),
      textInput("new_val1", "Recode to (should be 0)", value = "0", width = "100%"),
      selectInput("old_val2", "Original Value 2", choices = choices, 
                  selected = ifelse(length(choices) > 1, choices[2], choices[1]), width = "100%"),
      textInput("new_val2", "Recode to (should be 1)", value = "1", width = "100%")
    )
  })
  
  # Get color based on selection
  get_color_palette <- reactive({
    switch(input$color_palette,
           "blue" = c("#3498db", "#2980b9", "#1f618d"),
           "red" = c("#e74c3c", "#c0392b", "#a93226"),
           "green" = c("#2ecc71", "#27ae60", "#229954"),
           "purple" = c("#9b59b6", "#8e44ad", "#7d3c98"),
           "custom" = c(input$custom_color, 
                        colorRampPalette(c(input$custom_color, "white"))(3)[2],
                        colorRampPalette(c(input$custom_color, "white"))(3)[3])
    )
  })
  
  # Safe metric extraction function with enhanced error handling
  safe_extract_metrics <- function(conf_mat) {
    metrics <- list()
    
    # Initialize with default values
    metrics$sensitivity <- "N/A"
    metrics$specificity <- "N/A"
    metrics$accuracy <- "N/A"
    
    tryCatch({
      # Method 1: Extract from confusionMatrix byClass
      if (!is.null(conf_mat$byClass)) {
        if ("Sensitivity" %in% names(conf_mat$byClass) && 
            !is.na(conf_mat$byClass["Sensitivity"])) {
          metrics$sensitivity <- round(conf_mat$byClass["Sensitivity"], 4)
        }
        if ("Specificity" %in% names(conf_mat$byClass) && 
            !is.na(conf_mat$byClass["Specificity"])) {
          metrics$specificity <- round(conf_mat$byClass["Specificity"], 4)
        }
      }
      
      # Method 2: Extract from confusionMatrix overall
      if (!is.null(conf_mat$overall)) {
        if ("Accuracy" %in% names(conf_mat$overall) && 
            !is.na(conf_mat$overall["Accuracy"])) {
          metrics$accuracy <- round(conf_mat$overall["Accuracy"], 4)
        }
      }
      
      # Method 3: Manual calculation from confusion matrix table
      if (!is.null(conf_mat$table) && 
          all(dim(conf_mat$table) == c(2, 2))) {
        tn <- conf_mat$table[1, 1]
        fp <- conf_mat$table[1, 2]
        fn <- conf_mat$table[2, 1]
        tp <- conf_mat$table[2, 2]
        
        # Calculate sensitivity (recall)
        if ((tp + fn) > 0) {
          metrics$sensitivity <- round(tp / (tp + fn), 4)
        }
        
        # Calculate specificity
        if ((tn + fp) > 0) {
          metrics$specificity <- round(tn / (tn + fp), 4)
        }
        
        # Calculate accuracy
        total <- sum(conf_mat$table)
        if (total > 0) {
          correct <- sum(diag(conf_mat$table))
          metrics$accuracy <- round(correct / total, 4)
        }
      }
      
    }, error = function(e) {
      # If any error occurs, use default "N/A" values
      message("Error extracting metrics: ", e$message)
    })
    
    return(metrics)
  }
  
  # Enhanced analysis_result function with better error handling
  analysis_result <- eventReactive(input$generate, {
    tryCatch({
      df <- data_reactive()
      
      # Validate data exists
      if (is.null(df)) {
        stop("No data available. Please upload a file first.")
      }
      
      req(input$outcome, input$predictors)
      
      # Validate that variables exist in dataframe
      if (!input$outcome %in% names(df)) {
        stop("Selected outcome variable '", input$outcome, "' not found in dataset")
      }
      
      missing_predictors <- input$predictors[!input$predictors %in% names(df)]
      if (length(missing_predictors) > 0) {
        stop("The following predictor variables not found in dataset: ", 
             paste(missing_predictors, collapse = ", "))
      }
      
      # Apply recoding if specified
      if (!is.null(input$old_val1) && input$new_val1 != "" &&
          !is.null(input$old_val2) && input$new_val2 != "") {
        df[[input$outcome]] <- as.character(df[[input$outcome]])
        df[[input$outcome]][df[[input$outcome]] == input$old_val1] <- input$new_val1
        df[[input$outcome]][df[[input$outcome]] == input$old_val2] <- input$new_val2
      }
      
      # Select only needed variables and remove NAs
      all_vars <- c(input$outcome, input$predictors)
      df <- df[, all_vars, drop = FALSE]
      df <- na.omit(df)  # Omit rows with NA values
      
      # Check if we have data after removing NAs
      if (nrow(df) == 0) {
        stop("No data remaining after removing missing values. Please check your data.")
      }
      
      # Convert outcome to numeric and validate
      df[[input$outcome]] <- as.numeric(as.character(df[[input$outcome]]))
      unique_vals <- unique(na.omit(df[[input$outcome]]))
      
      # Ensure binary outcome
      if (!all(unique_vals %in% c(0, 1))) {
        stop("Outcome must be binary (0 and 1) after recoding. Current values: ", 
             paste(unique_vals, collapse = ", "), 
             ". Please use the recoding options above.")
      }
      
      # Check if we have both classes
      if (length(unique_vals) < 2) {
        stop("Outcome variable has only one class. Both classes (0 and 1) are required for ROC analysis.")
      }
      
      # Convert to factor with explicit levels
      df[[input$outcome]] <- factor(df[[input$outcome]], levels = c(0, 1))
      
      # Build model - handle single predictor case
      if (length(input$predictors) == 1) {
        formula <- as.formula(paste(input$outcome, "~", input$predictors))
      } else {
        formula <- as.formula(paste(input$outcome, "~", paste(input$predictors, collapse = "+")))
      }
      
      model <- tryCatch({
        glm(formula, data = df, family = "binomial")
      }, error = function(e) {
        stop("Model fitting failed: ", e$message)
      })
      
      # Generate predictions
      probs <- predict(model, type = "response")
      pred_class <- factor(ifelse(probs >= 0.5, 1, 0), levels = c(0, 1))
      true_class <- df[[input$outcome]]
      
      # Remove any remaining NAs
      keep <- complete.cases(probs, true_class)
      probs <- probs[keep]
      pred_class <- pred_class[keep]
      true_class <- true_class[keep]
      
      # Check if we have enough data
      if (length(probs) < 2) {
        stop("Insufficient data after removing missing values")
      }
      
      # Calculate ROC
      roc_obj <- tryCatch({
        roc(true_class, probs)
      }, error = function(e) {
        stop("ROC calculation failed: ", e$message)
      })
      
      # Generate AUC
      auc_val <- auc(roc_obj)
      
      # Calculate confusion matrix with robust error handling
      conf_mat <- tryCatch({
        # Ensure both vectors have the same levels
        if (!all(levels(pred_class) %in% levels(true_class))) {
          levels(pred_class) <- levels(true_class)
        }
        confusionMatrix(pred_class, true_class, positive = "1")
      }, error = function(e) {
        # Fallback: create basic confusion matrix manually
        tbl <- table(Predicted = pred_class, Actual = true_class)
        # Ensure the table has 2x2 dimensions
        if (all(dim(tbl) == c(2, 2))) {
          list(table = tbl)
        } else {
          # Create a dummy 2x2 table if dimensions don't match
          dummy_table <- matrix(c(1, 0, 0, 1), nrow = 2, 
                                dimnames = list(Predicted = c("0", "1"), 
                                                Actual = c("0", "1")))
          list(table = dummy_table)
        }
      })
      
      # Extract metrics safely
      metrics <- safe_extract_metrics(conf_mat)
      
      # Prepare ROC data for plotting
      roc_df <- data.frame(
        TPR = roc_obj$sensitivities,
        FPR = 1 - roc_obj$specificities
      )
      roc_df <- roc_df[complete.cases(roc_df), ]  # Remove any NAs
      roc_df <- roc_df[order(roc_df$FPR, roc_df$TPR), ]
      
      # Check if we have valid ROC data for plotting
      if (nrow(roc_df) < 2) {
        stop("Insufficient data to create ROC curve")
      }
      
      # Get color palette
      colors <- get_color_palette()
      
      # Create advanced ggplot2 ROC plot
      p <- ggplot(roc_df, aes(x = FPR, y = TPR)) +
        # ROC curve line
        geom_line(color = colors[1], size = input$line_size, alpha = input$alpha_level) +
        # Area under curve
        geom_ribbon(aes(ymin = 0, ymax = TPR), fill = colors[1], alpha = 0.2) +
        # Diagonal reference line
        {if(input$show_diagonal) geom_abline(linetype = "dashed", color = "gray50", slope = 1, intercept = 0, size = 0.8)} +
        # AUC annotation
        {if(input$show_auc) annotate("text", x = 0.7, y = 0.2, 
                                     label = paste("AUC =", round(auc_val, 3)), 
                                     size = 5, fontface = "bold", color = colors[1])} +
        # Labels and title
        labs(title = input$graph_title, 
             x = input$x_label, 
             y = input$y_label) +
        # Professional theme
        theme_minimal(base_size = 14) +
        theme(
          plot.title = element_text(hjust = 0.5, face = "bold", size = 18, margin = margin(b = 20)),
          axis.title = element_text(face = "bold", size = 12),
          axis.text = element_text(size = 10),
          panel.grid.major = element_line(color = "grey90", size = 0.2),
          panel.grid.minor = element_blank(),
          panel.background = element_rect(fill = "white", color = NA),
          plot.background = element_rect(fill = "white", color = NA),
          legend.position = "none"
        ) +
        # Equal coordinates and limits
        coord_equal() +
        scale_x_continuous(limits = c(0, 1), expand = c(0, 0)) +
        scale_y_continuous(limits = c(0, 1), expand = c(0, 0))
      
      list(
        roc_obj = roc_obj,
        auc = auc_val,
        accuracy = metrics$accuracy,
        sensitivity = metrics$sensitivity,
        specificity = metrics$specificity,
        conf_mat = conf_mat,
        plot = p,
        success = TRUE
      )
    }, error = function(e) {
      return(list(error = TRUE, message = e$message, success = FALSE))
    })
  })
  
  # ROC Plot Output - UPDATED to use ggplot2 instead of plotly
  output$roc_plot <- renderPlot({
    res <- analysis_result()
    
    # Check if result is NULL or has error
    if (is.null(res)) {
      # Create empty plot with message
      ggplot() +
        annotate("text", x = 0.5, y = 0.5, 
                 label = "Please upload data and click 'Generate ROC Curve'", 
                 size = 6, fontface = "bold") +
        theme_void() +
        theme(plot.background = element_rect(fill = "white", color = NA))
    } else if (!is.null(res$error) && res$error) {
      # Create error plot
      error_msg <- if (!is.null(res$message)) res$message else "Analysis failed"
      ggplot() +
        annotate("text", x = 0.5, y = 0.5, 
                 label = paste("Error:", error_msg), 
                 size = 5, fontface = "bold", color = "red") +
        theme_void() +
        theme(plot.background = element_rect(fill = "white", color = NA))
    } else if (!is.null(res$success) && !res$success) {
      # Create failure plot
      error_msg <- if (!is.null(res$message)) res$message else "Analysis failed"
      ggplot() +
        annotate("text", x = 0.5, y = 0.5, 
                 label = paste("Error:", error_msg), 
                 size = 5, fontface = "bold", color = "red") +
        theme_void() +
        theme(plot.background = element_rect(fill = "white", color = NA))
    } else if (is.null(res$plot)) {
      # Create no plot available message
      ggplot() +
        annotate("text", x = 0.5, y = 0.5, 
                 label = "No plot available", 
                 size = 6, fontface = "bold") +
        theme_void() +
        theme(plot.background = element_rect(fill = "white", color = NA))
    } else {
      # Return the actual ROC plot
      res$plot
    }
  }, height = 500)
  
  # Performance Metrics Table
  output$metrics_table <- renderDT({
    res <- analysis_result()
    
    if (is.null(res) || !is.null(res$error) || !res$success) {
      return(datatable(data.frame(Message = "No metrics available. Please run analysis first."),
                       options = list(dom = 't'), rownames = FALSE))
    }
    
    metrics_df <- data.frame(
      Metric = c("AUC", "Accuracy", "Sensitivity", "Specificity"),
      Value = c(
        round(as.numeric(res$auc), 4),
        ifelse(is.numeric(res$accuracy), round(res$accuracy, 4), res$accuracy),
        ifelse(is.numeric(res$sensitivity), round(res$sensitivity, 4), res$sensitivity),
        ifelse(is.numeric(res$specificity), round(res$specificity, 4), res$specificity)
      )
    )
    
    datatable(metrics_df, 
              options = list(dom = 't', ordering = FALSE), 
              rownames = FALSE) %>%
      formatStyle('Value', fontWeight = 'bold')
  })
  
  # Enhanced confusion matrix table rendering
  output$conf_matrix_table <- renderDT({
    res <- analysis_result()
    if (is.null(res) || !is.null(res$error) || !res$success) {
      return(datatable(data.frame(Message = "No confusion matrix available. Please run analysis first."),
                       options = list(dom = 't'), rownames = FALSE))
    }
    
    # Safely extract confusion matrix table with enhanced error handling
    cm_table <- tryCatch({
      if (!is.null(res$conf_mat$table) && 
          all(dim(res$conf_mat$table) == c(2, 2))) {
        # Proper 2x2 confusion matrix
        tbl_df <- as.data.frame.matrix(res$conf_mat$table)
        colnames(tbl_df) <- c("Actual 0", "Actual 1")
        rownames(tbl_df) <- c("Predicted 0", "Predicted 1")
        tbl_df
      } else {
        # Create empty table with proper structure
        empty_df <- data.frame(`Actual 0` = c("N/A", "N/A"), 
                               `Actual 1` = c("N/A", "N/A"),
                               row.names = c("Predicted 0", "Predicted 1"))
        colnames(empty_df) <- c("Actual 0", "Actual 1")
        empty_df
      }
    }, error = function(e) {
      # Fallback table
      fallback_df <- data.frame(`Actual 0` = c("Error", ""), 
                                `Actual 1` = c("", "Error"),
                                row.names = c("Predicted 0", "Predicted 1"))
      colnames(fallback_df) <- c("Actual 0", "Actual 1")
      fallback_df
    })
    
    datatable(cm_table, 
              options = list(dom = 't', ordering = FALSE), 
              rownames = TRUE) %>%
      formatStyle(names(cm_table), fontWeight = 'bold') %>%
      formatStyle(colnames(cm_table)[1], color = '#e74c3c') %>%
      formatStyle(colnames(cm_table)[2], color = '#2ecc71')
  })
  
  # Download handlers - UPDATED to ensure AUC is labeled on downloaded PNG
  output$download_plot <- downloadHandler(
    filename = function() {
      paste0("ROC_Curve_", Sys.Date(), ".png")
    },
    content = function(file) {
      res <- analysis_result()
      if (is.null(res) || !res$success) {
        stop("Cannot download plot: Analysis failed or not run")
      }
      
      # Save high-quality PNG with AUC label
      ggsave(file, 
             plot = res$plot, 
             device = "png", 
             width = 10, 
             height = 8, 
             dpi = 300, 
             bg = "white")
    }
  )
  
  output$download_word <- downloadHandler(
    filename = function() {
      paste0("ROC_Report_", Sys.Date(), ".docx")
    },
    content = function(file) {
      res <- analysis_result()
      if (is.null(res) || !res$success) {
        stop("Cannot generate report: Analysis failed or not run")
      }
      
      tmp_img <- tempfile(fileext = ".png")
      ggsave(tmp_img, plot = res$plot, width = 7, height = 5, dpi = 300)
      
      # Create simple report without flextable
      doc <- read_docx() %>%
        body_add_par("ROC Curve Analysis Report", style = "heading 1") %>%
        body_add_par(paste("Generated on", format(Sys.Date(), "%B %d, %Y")), style = "Normal") %>%
        body_add_break() %>%
        body_add_img(tmp_img, width = 6, height = 4.5) %>%
        body_add_par("Performance Metrics", style = "heading 2")
      
      # Add metrics as simple text
      metrics_text <- paste(
        "AUC:", round(as.numeric(res$auc), 4), "\n",
        "Accuracy:", ifelse(is.numeric(res$accuracy), round(res$accuracy, 4), res$accuracy), "\n",
        "Sensitivity:", ifelse(is.numeric(res$sensitivity), round(res$sensitivity, 4), res$sensitivity), "\n",
        "Specificity:", ifelse(is.numeric(res$specificity), round(res$specificity, 4), res$specificity)
      )
      
      doc <- doc %>%
        body_add_par(metrics_text, style = "Normal") %>%
        body_add_par("Analysis Details", style = "heading 2") %>%
        body_add_par(paste("Outcome variable:", input$outcome), style = "Normal") %>%
        body_add_par(paste("Predictor variables:", paste(input$predictors, collapse = ", ")), style = "Normal") %>%
        body_add_par(paste("Observations:", nrow(na.omit(data_reactive()[, c(input$outcome, input$predictors)]))), style = "Normal") %>%
        body_add_break() %>%
        body_add_par("ROC Curve Builder - For research and educational use", style = "Normal")
      
      print(doc, target = file)
      
      # Clean up temporary file
      unlink(tmp_img)
    }
  )
  
  # Navigation button
  observeEvent(input$goto_analysis, {
    updateNavbarPage(session, "main_nav", selected = "Build ROC Curve")
  })
}

shinyApp(ui, server)
