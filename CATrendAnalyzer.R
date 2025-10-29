library(shiny)
library(shinydashboard)
library(DT)
library(haven)
library(readxl)
library(ggplot2)
library(officer)
library(flextable)
library(rmarkdown)
library(tidyverse)
library(sortable)

# Helper function to convert haven_labelled variables to factors with labels
convert_labelled_to_factor <- function(df) {
  for (col_name in names(df)) {
    if (inherits(df[[col_name]], "haven_labelled")) {
      # Extract value labels
      val_labels <- attr(df[[col_name]], "labels")
      
      if (!is.null(val_labels) && length(val_labels) > 0) {
        # Convert to factor with labels
        df[[col_name]] <- as_factor(df[[col_name]])
      } else {
        # If no labels, convert to regular factor or character
        df[[col_name]] <- as.character(df[[col_name]])
      }
    }
  }
  return(df)
}

# Define UI
ui <- dashboardPage(
  skin = "blue",
  dashboardHeader(
    title = span(
      icon("chart-line"), 
      "CATrend Analyzer",
      style = "font-weight: bold; font-size: 20px; color: #FFFFFF;"
    ),
    titleWidth = 350
  ),
  
  dashboardSidebar(
    width = 250,
    sidebarMenu(
      id = "tabs",
      menuItem("Home", tabName = "home", icon = icon("home")),
      menuItem("Data Upload", tabName = "upload", icon = icon("upload")),
      menuItem("Data Cleaning", tabName = "cleaning", icon = icon("broom")),
      menuItem("Variable Selection", tabName = "variables", icon = icon("sliders")),
      menuItem("Analysis", tabName = "analysis", icon = icon("chart-bar")),
      menuItem("Results", tabName = "results", icon = icon("file-download"))
    )
  ),
  
  dashboardBody(
    tags$head(
      tags$style(HTML("
        /* Fix sidebar and header to be static */
        .wrapper {
          height: 100vh !important;
          overflow: hidden !important;
        }
        
        .main-sidebar {
          position: fixed !important;
          height: 100vh !important;
          overflow-y: auto !important;
        }
        
        .content-wrapper {
          margin-left: 250px !important;
          height: 100vh !important;
          overflow-y: auto !important;
        }
        
        .main-header .navbar {
          margin-left: 250px !important;
        }
        
        .main-header .logo {
          position: fixed !important;
          width: 250px !important;
        }
        
        /* Ensure content area scrolls properly */
        .content {
          padding: 20px;
          min-height: calc(100vh - 100px);
        }
        
        /* WHO-inspired color scheme - Blue theme */
        .skin-blue .main-header .logo {
          background-color: #0093D0;
          color: #ffffff;
          font-weight: bold;
        }
        
        .skin-blue .main-header .logo:hover {
          background-color: #007BB8;
        }
        
        .skin-blue .main-header .navbar {
          background-color: #0093D0;
        }
        
        .skin-blue .main-sidebar {
          background-color: #1A5276;
        }
        
        .skin-blue .main-sidebar .sidebar .sidebar-menu .active a {
          background-color: #2E86AB;
          border-left-color: #0093D0;
        }
        
        .skin-blue .main-sidebar .sidebar .sidebar-menu a {
          color: #E8F4F8;
          border-left: 3px solid transparent;
        }
        
        .skin-blue .main-sidebar .sidebar .sidebar-menu a:hover {
          background-color: #2E86AB;
          color: #ffffff;
          border-left-color: #0093D0;
        }
        
        /* Main styling */
        .content-wrapper, .right-side {
          background-color: #f8f9fa;
        }
        
        .box {
          box-shadow: 0 4px 6px rgba(0,0,0,0.1);
          border-radius: 8px;
          border-top: 3px solid #0093D0;
        }
        
        .box.box-primary {
          border-top-color: #0093D0;
        }
        
        .box.box-info {
          border-top-color: #2E86AB;
        }
        
        .box.box-success {
          border-top-color: #27AE60;
        }
        
        .box.box-warning {
          border-top-color: #F39C12;
        }
        
        .rank-list-container {
          min-height: 100px;
          border: 1px solid #ddd;
          padding: 10px;
          margin-bottom: 10px;
          border-radius: 5px;
          background-color: #f8f9fa;
        }
        
        .rank-list-item {
          padding: 8px 12px;
          margin: 2px;
          background-color: white;
          border: 1px solid #dee2e6;
          cursor: move;
          border-radius: 4px;
          transition: all 0.2s ease;
        }
        
        .rank-list-item:hover {
          background-color: #e3f2fd;
          border-color: #2196f3;
        }
        
        .missing-data {
          background-color: #fff3cd;
          padding: 15px;
          border-radius: 5px;
          border: 1px solid #ffeaa7;
        }
        
        .btn {
          border-radius: 4px;
          font-weight: 500;
        }
        
        .btn-primary {
          background-color: #0093D0;
          border-color: #007BB8;
        }
        
        .btn-success {
          background-color: #27AE60;
          border-color: #219653;
        }
        
        .btn-warning {
          background-color: #F39C12;
          border-color: #E67E22;
        }
        
        /* Instruction highlight */
        .instruction-highlight {
          background-color: #FFF9C4;
          border-left: 4px solid #F39C12;
          padding: 12px 15px;
          margin: 10px 0;
          border-radius: 4px;
          font-weight: 500;
          color: #7D6608;
        }
        
        /* Home page styling */
        .home-header {
          background: linear-gradient(135deg, #0093D0 0%, #1A5276 100%);
          color: white;
          padding: 40px 20px;
          text-align: center;
          border-radius: 8px;
          margin-bottom: 30px;
        }
        
        .home-header h1 {
          font-size: 2.5em;
          margin-bottom: 10px;
          font-weight: 700;
        }
        
        .home-header p {
          font-size: 1.2em;
          opacity: 0.9;
        }
        
        .feature-card {
          background: white;
          padding: 25px;
          border-radius: 8px;
          box-shadow: 0 2px 4px rgba(0,0,0,0.1);
          margin-bottom: 20px;
          border-left: 4px solid #0093D0;
          transition: transform 0.2s ease;
        }
        
        .feature-card:hover {
          transform: translateY(-2px);
          box-shadow: 0 4px 8px rgba(0,0,0,0.15);
        }
        
        .feature-card h3 {
          color: #2c3e50;
          margin-top: 0;
          font-weight: 600;
        }
        
        .feature-card p {
          color: #6c757d;
          line-height: 1.6;
        }
        
        .info-section {
          background: white;
          padding: 30px;
          border-radius: 8px;
          box-shadow: 0 2px 4px rgba(0,0,0,0.1);
          margin-bottom: 20px;
        }
        
        .info-section h2 {
          color: #2c3e50;
          border-bottom: 2px solid #e9ecef;
          padding-bottom: 10px;
          margin-bottom: 20px;
          font-weight: 600;
        }
        
        .contact-info {
          background: linear-gradient(135deg, #0093D0 0%, #1A5276 100%);
          color: white;
          padding: 25px;
          border-radius: 8px;
          text-align: center;
        }
        
        .contact-info h3 {
          margin-top: 0;
          font-weight: 600;
        }
        
        .contact-info a {
          color: white;
          text-decoration: underline;
          font-weight: 500;
        }
        
        /* Analysis tab enhancements */
        .analysis-controls {
          background: white;
          padding: 20px;
          border-radius: 8px;
          margin-bottom: 20px;
          border-left: 4px solid #27AE60;
        }
        
        .results-panel {
          background: white;
          padding: 25px;
          border-radius: 8px;
          box-shadow: 0 2px 4px rgba(0,0,0,0.1);
          margin-bottom: 20px;
          border-left: 4px solid #0093D0;
        }
        
        .results-panel h3 {
          color: #2c3e50;
          border-bottom: 2px solid #e9ecef;
          padding-bottom: 10px;
          margin-bottom: 20px;
          font-weight: 600;
        }
      "))
    ),
    
    tabItems(
      # Home Tab
      tabItem(tabName = "home",
              fluidRow(
                column(12,
                       div(class = "home-header",
                           h1("CATrend Analyzer"),
                           p("An Open Source Tool for Performing Cochran-Armitage Trend Test")
                       )
                )
              ),
              
              fluidRow(
                column(8,
                       div(class = "info-section",
                           h2("About This Application"),
                           p("CATrend Analyzer performs the Cochran-Armitage test for trend, a specialized statistical method for detecting trends in proportions across ordered categorical variables. This test is widely used in clinical trials, epidemiology, and public health research, particularly when assessing whether the probability of a binary outcome (e.g., disease occurrence, treatment response) changes consistently across increasing levels of an exposure or dose."),
                           
                           h3("What the App Does"),
                           p("The CATrend Analyzer provides:"),
                           tags$ul(
                             tags$li("Data import from multiple formats (CSV, Excel, Stata, SPSS)"),
                             tags$li("Advanced data cleaning and missing value handling"),
                             tags$li("Interactive variable selection and level ordering"),
                             tags$li("Accurate Cochran-Armitage test implementation"),
                             tags$li("Professional visualization and reporting"),
                             tags$li("Exportable results in word format")
                           ),
                           
                           h3("Statistical Methodology"),
                           p("The application implements the exact Cochran-Armitage test formula:"),
                           withMathJax(),
                           helpText("$$Z = \\frac{\\sum_{j=1}^k y_j(x_j - \\bar{x})}{\\sqrt{\\bar{p}(1-\\bar{p})\\sum_{j=1}^k n_j(x_j - \\bar{x})^2}}$$"),
                           p("Where:"),
                           tags$ul(
                             tags$li("\\(x_j\\) are the scores for the ordered groups"),
                             tags$li("\\(y_j\\) are the success counts in each group"),
                             tags$li("\\(n_j\\) are the total counts in each group"),
                             tags$li("\\(\\bar{x}\\) is the weighted mean of scores"),
                             tags$li("\\(\\bar{p}\\) is the overall proportion of successes")
                           ),
                           
                           h3("Key Metrics"),
                           p("The analysis provides:"),
                           tags$ul(
                             tags$li("Z-statistic and Chi-square values"),
                             tags$li("Exact p-values for trend detection"),
                             tags$li("Contingency tables with proportions"),
                             tags$li("Visual trend analysis through stacked bar plots"),
                             tags$li("Comprehensive data quality assessments")
                           )
                       )
                ),
                
                column(4,
                       div(class = "feature-card",
                           h3("Developer Information"),
                           p("Developed by:"),
                           tags$h4("Mudasir Mohammed Ibrahim"),
                           tags$p("Registered Nurse"),
                           tags$p("Email: ", tags$a(href="mailto:mudassiribrahim30@gmail.com", "mudassiribrahim30@gmail.com"))
                       ),
                       
                       div(class = "feature-card",
                           h3("Application Features"),
                           tags$ul(
                             tags$li("Multi-format data support"),
                             tags$li("Interactive data cleaning"),
                             tags$li("Drag-and-drop variable ordering"),
                             tags$li("Professional reporting")
                           )
                       ),
                       
                       div(class = "contact-info",
                           h3("Support & Contact"),
                           p("For any suggestions, bug reports, or problems with the application, please contact me at:"),
                           tags$h4(tags$a(href="mailto:mudassiribrahim30@gmail.com", "mudassiribrahim30@gmail.com")),
                           p("I welcome your feedback and will respond promptly to all inquiries.")
                       )
                )
              )
      ),
      
      # Data Upload Tab
      tabItem(tabName = "upload",
              fluidRow(
                box(
                  title = "Upload Data File", status = "primary", solidHeader = TRUE, width = 12,
                  radioButtons("file_type", "Select File Type:",
                               choices = c("CSV" = "csv", "Excel" = "excel", 
                                           "Stata" = "stata", "SPSS" = "spss"),
                               selected = "csv", inline = TRUE),
                  
                  conditionalPanel(
                    condition = "input.file_type == 'csv'",
                    fileInput("file_csv", "Choose CSV File",
                              accept = c("text/csv", "text/comma-separated-values",
                                         "text/plain", ".csv")),
                    checkboxInput("header_csv", "Header", TRUE),
                    radioButtons("sep_csv", "Separator",
                                 choices = c(Comma = ",", Semicolon = ";", Tab = "\t"),
                                 selected = ",", inline = TRUE),
                    radioButtons("quote_csv", "Quote",
                                 choices = c(None = "", "Double Quote" = '"', "Single Quote" = "'"),
                                 selected = '"', inline = TRUE)
                  ),
                  
                  conditionalPanel(
                    condition = "input.file_type == 'excel'",
                    fileInput("file_excel", "Choose Excel File",
                              accept = c(".xls", ".xlsx")),
                    numericInput("sheet_excel", "Sheet Number", value = 1, min = 1),
                    textInput("range_excel", "Cell Range (optional)", placeholder = "A1:Z100")
                  ),
                  
                  conditionalPanel(
                    condition = "input.file_type == 'stata'",
                    fileInput("file_stata", "Choose Stata File",
                              accept = c(".dta"))
                  ),
                  
                  conditionalPanel(
                    condition = "input.file_type == 'spss'",
                    fileInput("file_spss", "Choose SPSS File",
                              accept = c(".sav"))
                  )
                )
              ),
              
              fluidRow(
                box(
                  title = "Data Preview", status = "info", solidHeader = TRUE, width = 12,
                  DTOutput("data_preview")
                )
              )
      ),
      
      # Data Cleaning Tab
      tabItem(tabName = "cleaning",
              fluidRow(
                box(
                  title = "Missing Data Overview", status = "warning", solidHeader = TRUE, width = 12,
                  htmlOutput("missing_data_summary"),
                  actionButton("clean_data", "Remove Rows with Missing Values in Selected Variables", 
                               class = "btn-warning", icon = icon("broom"))
                )
              ),
              
              fluidRow(
                box(
                  title = "Data Quality Check", status = "info", solidHeader = TRUE, width = 6,
                  verbatimTextOutput("data_quality")
                ),
                box(
                  title = "Cleaned Data Preview", status = "success", solidHeader = TRUE, width = 6,
                  DTOutput("cleaned_data_preview")
                )
              )
      ),
      
      # Variable Selection Tab
      tabItem(tabName = "variables",
              fluidRow(
                box(
                  title = "Variable Selection", status = "primary", solidHeader = TRUE, width = 6,
                  uiOutput("dependent_var_ui"),
                  uiOutput("independent_var_ui"),
                  div(class = "instruction-highlight",
                      icon("exclamation-triangle"),
                      "IMPORTANT: When ordering independent variable levels, drag them from LOWEST to HIGHEST (e.g., Certificate -> Bachelor's -> Master's -> PhD for education level). After that, click on Update Level Ordering."
                  )
                ),
                
                box(
                  title = "Advanced Options", status = "warning", solidHeader = TRUE, width = 6,
                  uiOutput("dependent_levels_ui"),
                  uiOutput("independent_levels_ui"),
                  actionButton("update_levels", "Update Level Ordering", 
                               class = "btn-warning")
                )
              ),
              
              fluidRow(
                box(
                  title = "Current Variable Information", status = "info", solidHeader = TRUE, width = 12,
                  verbatimTextOutput("var_info")
                )
              )
      ),
      
      # Analysis Tab
      tabItem(tabName = "analysis",
              fluidRow(
                box(
                  title = "Analysis Controls", status = "primary", solidHeader = TRUE, width = 12,
                  div(class = "analysis-controls",
                      radioButtons("test_type", "Test Type:",
                                   choices = c("Two-sided" = "two.sided", 
                                               "One-sided (increasing)" = "one.sided.increasing",
                                               "One-sided (decreasing)" = "one.sided.decreasing"),
                                   selected = "two.sided", inline = TRUE),
                      checkboxInput("continuity_correction", "Apply Continuity Correction", value = FALSE),
                      actionButton("analyze", "Run Cochran-Armitage Test", 
                                   class = "btn-primary", icon = icon("play")),
                      helpText("Note: Analysis will use the cleaned data from the Data Cleaning tab")
                  )
                )
              ),
              
              # Cochran-Armitage Test Results moved up (second thing users see)
              fluidRow(
                box(
                  title = "Cochran-Armitage Test Results", status = "success", solidHeader = TRUE, width = 12,
                  div(class = "results-panel",
                      verbatimTextOutput("test_results")
                  )
                )
              ),
              
              # Stacked Bar Plot moved down (third thing users see)
              fluidRow(
                box(
                  title = "Stacked Bar Plot", status = "info", solidHeader = TRUE, width = 12,
                  plotOutput("bar_plot", height = "500px")
                )
              )
      ),
      
      # Results Download Tab
      tabItem(tabName = "results",
              fluidRow(
                box(
                  title = "Download Results", status = "primary", solidHeader = TRUE, width = 12,
                  radioButtons("report_format", "Select Report Format:",
                               choices = c("Word Document" = "word"),
                               selected = "word", inline = TRUE),
                  conditionalPanel(
                    condition = "output.analysis_done",
                    downloadButton("download_report", "Download Full Report",
                                   class = "btn-success")
                  ),
                  conditionalPanel(
                    condition = "!output.analysis_done",
                    div(class = "instruction-highlight",
                        icon("info-circle"),
                        "Please run the analysis first in the Analysis tab to generate results for download."
                    )
                  )
                )
              ),
              
              fluidRow(
                box(
                  title = "Report Preview", status = "info", solidHeader = TRUE, width = 12,
                  htmlOutput("report_preview")
                )
              )
      )
    )
  )
)

# Define Server
server <- function(input, output, session) {
  
  # Reactive values
  data_loaded <- reactiveVal(NULL)
  data_cleaned <- reactiveVal(NULL)
  dep_level_order <- reactiveVal(NULL)
  ind_level_order <- reactiveVal(NULL)
  analysis_done <- reactiveVal(FALSE)
  
  # Track if analysis has been completed
  output$analysis_done <- reactive({
    analysis_done()
  })
  outputOptions(output, "analysis_done", suspendWhenHidden = FALSE)
  
  # Load data based on file type - UPDATED FOR SPSS AND STATA
  observe({
    tryCatch({
      req(input$file_type)
      
      if (input$file_type == "csv" && !is.null(input$file_csv)) {
        df <- read.csv(input$file_csv$datapath,
                       header = input$header_csv,
                       sep = input$sep_csv,
                       quote = input$quote_csv,
                       stringsAsFactors = FALSE,
                       na.strings = c("", " ", "NA", "N/A", "NULL", "null", "NaN"))
        data_loaded(df)
        data_cleaned(df)
        
      } else if (input$file_type == "excel" && !is.null(input$file_excel)) {
        if (!is.null(input$range_excel) && input$range_excel != "") {
          df <- read_excel(input$file_excel$datapath, 
                           sheet = input$sheet_excel,
                           range = input$range_excel)
        } else {
          df <- read_excel(input$file_excel$datapath, 
                           sheet = input$sheet_excel)
        }
        data_loaded(as.data.frame(df))
        data_cleaned(as.data.frame(df))
        
      } else if (input$file_type == "stata" && !is.null(input$file_stata)) {
        df <- read_stata(input$file_stata$datapath)
        # Convert labelled variables to factors with labels
        df <- convert_labelled_to_factor(df)
        data_loaded(as.data.frame(df))
        data_cleaned(as.data.frame(df))
        
      } else if (input$file_type == "spss" && !is.null(input$file_spss)) {
        df <- read_sav(input$file_spss$datapath)
        # Convert labelled variables to factors with labels
        df <- convert_labelled_to_factor(df)
        data_loaded(as.data.frame(df))
        data_cleaned(as.data.frame(df))
      }
    }, error = function(e) {
      showNotification(paste("Error loading file:", e$message), type = "error")
    })
  })
  
  # Data preview
  output$data_preview <- renderDT({
    req(data_loaded())
    datatable(data_loaded(), 
              options = list(scrollX = TRUE, pageLength = 5),
              caption = "Original Data Preview")
  })
  
  # Missing data summary
  output$missing_data_summary <- renderUI({
    req(data_loaded())
    
    df <- data_loaded()
    missing_count <- colSums(is.na(df))
    missing_pct <- round((missing_count / nrow(df)) * 100, 2)
    total_missing <- sum(is.na(df))
    total_missing_pct <- round((total_missing / (nrow(df) * ncol(df))) * 100, 2)
    
    tagList(
      h4("Missing Data Summary"),
      p("Total missing values: ", total_missing, " (", total_missing_pct, "% of all cells)"),
      p("Rows with any missing values: ", sum(!complete.cases(df)), " (", 
        round((sum(!complete.cases(df)) / nrow(df)) * 100, 2), "% of rows)"),
      br(),
      h5("Missing values by column:"),
      renderTable({
        data.frame(
          Variable = names(df),
          Missing_Count = missing_count,
          Missing_Percent = paste0(missing_pct, "%"),
          stringsAsFactors = FALSE
        )
      }, striped = TRUE, hover = TRUE, width = "100%")
    )
  })
  
  # Data quality check
  output$data_quality <- renderPrint({
    req(data_loaded())
    
    df <- data_loaded()
    cat("DATA QUALITY REPORT\n")
    cat("===================\n\n")
    cat("Dataset Dimensions:", nrow(df), "rows ×", ncol(df), "columns\n")
    cat("Complete cases (no missing values):", sum(complete.cases(df)), "rows\n")
    cat("Rows with missing values:", sum(!complete.cases(df)), "rows\n")
    cat("Total missing values:", sum(is.na(df)), "\n")
    cat("Percentage of missing data:", round((sum(is.na(df)) / (nrow(df) * ncol(df))) * 100, 2), "%\n\n")
    
    cat("VARIABLE TYPES:\n")
    for (col in names(df)) {
      cat("- ", col, ": ", class(df[[col]]), 
          " (", length(unique(df[[col]])), " unique values)\n", sep = "")
    }
  })
  
  # Clean data
  observeEvent(input$clean_data, {
    req(data_loaded())
    
    df <- data_loaded()
    # Remove rows with any missing values
    df_clean <- df[complete.cases(df), ]
    
    if (nrow(df_clean) == 0) {
      showNotification("Warning: Removing all missing values would result in empty dataset!", 
                       type = "warning")
    } else {
      data_cleaned(df_clean)
      showNotification(paste("Data cleaned! Removed", nrow(df) - nrow(df_clean), 
                             "rows with missing values."), type = "message")
    }
  })
  
  # Cleaned data preview
  output$cleaned_data_preview <- renderDT({
    req(data_cleaned())
    datatable(data_cleaned(), 
              options = list(scrollX = TRUE, pageLength = 5),
              caption = "Cleaned Data Preview (Missing Values Removed)")
  })
  
  # Variable selection UIs
  output$dependent_var_ui <- renderUI({
    req(data_cleaned())
    selectInput("dependent_var", "Select Dependent Variable (Categorical):",
                choices = names(data_cleaned()))
  })
  
  output$independent_var_ui <- renderUI({
    req(data_cleaned())
    selectInput("independent_var", "Select Independent Variable (Ordinal):",
                choices = names(data_cleaned()))
  })
  
  # Level ordering UIs using sortable
  output$dependent_levels_ui <- renderUI({
    req(data_cleaned(), input$dependent_var)
    
    var_levels <- unique(data_cleaned()[[input$dependent_var]])
    var_levels <- var_levels[!is.na(var_levels)]
    
    if(length(var_levels) > 0) {
      tagList(
        h5("Order Dependent Variable Levels:"),
        helpText("Drag to reorder levels (current order will be used for analysis)"),
        bucket_list(
          header = NULL,
          group_name = "dep_bucket",
          orientation = "vertical",
          add_rank_list(
            text = "Drag from here",
            labels = as.character(var_levels),
            input_id = "dep_source"
          ),
          add_rank_list(
            text = "To here (current order)",
            labels = NULL,
            input_id = "dep_target"
          )
        )
      )
    }
  })
  
  output$independent_levels_ui <- renderUI({
    req(data_cleaned(), input$independent_var)
    
    var_levels <- unique(data_cleaned()[[input$independent_var]])
    var_levels <- var_levels[!is.na(var_levels)]
    
    if(length(var_levels) > 0) {
      tagList(
        h5("Order Independent Variable Levels:"),
        helpText("Drag levels from LOWEST to HIGHEST (top = lowest, bottom = highest)"),
        bucket_list(
          header = NULL,
          group_name = "ind_bucket",
          orientation = "vertical",
          add_rank_list(
            text = "Drag from here",
            labels = as.character(var_levels),
            input_id = "ind_source"
          ),
          add_rank_list(
            text = "To here (current order: lowest to highest)",
            labels = NULL,
            input_id = "ind_target"
          )
        )
      )
    }
  })
  
  # Update level ordering
  observeEvent(input$update_levels, {
    if(!is.null(input$dep_target)) {
      dep_level_order(input$dep_target)
    }
    if(!is.null(input$ind_target)) {
      ind_level_order(input$ind_target)
    }
    showNotification("Level ordering updated!", type = "message")
  })
  
  # Initialize level orders when variables are selected
  observeEvent(input$dependent_var, {
    req(data_cleaned(), input$dependent_var)
    var_levels <- unique(data_cleaned()[[input$dependent_var]])
    var_levels <- var_levels[!is.na(var_levels)]
    dep_level_order(as.character(var_levels))
  })
  
  observeEvent(input$independent_var, {
    req(data_cleaned(), input$independent_var)
    var_levels <- unique(data_cleaned()[[input$independent_var]])
    var_levels <- var_levels[!is.na(var_levels)]
    ind_level_order(as.character(var_levels))
  })
  
  # Variable information
  output$var_info <- renderPrint({
    req(data_cleaned(), input$dependent_var, input$independent_var)
    
    cat("DATA OVERVIEW:\n")
    cat("Rows:", nrow(data_cleaned()), "\n")
    cat("Columns:", ncol(data_cleaned()), "\n")
    cat("Complete cases:", sum(complete.cases(data_cleaned())), "\n\n")
    
    cat("DEPENDENT VARIABLE:\n")
    cat("Name:", input$dependent_var, "\n")
    dep_levels <- unique(data_cleaned()[[input$dependent_var]])
    dep_levels <- dep_levels[!is.na(dep_levels)]
    current_dep_order <- if(!is.null(dep_level_order())) dep_level_order() else as.character(dep_levels)
    cat("Levels:", length(current_dep_order), "\n")
    cat("Current Order:", paste(current_dep_order, collapse = " -> "), "\n\n")
    
    cat("INDEPENDENT VARIABLE:\n")
    cat("Name:", input$independent_var, "\n")
    ind_levels <- unique(data_cleaned()[[input$independent_var]])
    ind_levels <- ind_levels[!is.na(ind_levels)]
    current_ind_order <- if(!is.null(ind_level_order())) ind_level_order() else as.character(ind_levels)
    cat("Levels:", length(current_ind_order), "\n")
    cat("Current Order:", paste(current_ind_order, collapse = " < "), "\n\n")
    
    # Check if we have enough data for analysis
    df_subset <- data_cleaned()[, c(input$dependent_var, input$independent_var)]
    complete_cases <- sum(complete.cases(df_subset))
    cat("ANALYSIS READINESS:\n")
    cat("Complete cases for selected variables:", complete_cases, "\n")
    
    if (complete_cases < 10) {
      cat("⚠️ WARNING: Very small sample size for analysis\n")
    } else if (complete_cases < 30) {
      cat("⚠️ NOTE: Small sample size for analysis\n")
    } else {
      cat("✓ Sufficient sample size for analysis\n")
    }
  })
  
  # ACCURATE Cochran-Armitage test implementation using the provided formula
  perform_cochran_armitage_accurate <- function(cont_table, test_type = "two.sided", continuity_correction = FALSE) {
    # Implementation following the exact theoretical background provided
    # Let xj be the numeric labels of the columns (j = 1, 2, …k)
    # nj = the number of elements in the jth column
    # yj = the value in the cell in the first row and jth column
    
    k <- ncol(cont_table)  # Number of columns (ordered groups)
    N <- sum(cont_table)   # Total sample size
    
    # Assign scores to independent variable levels (1, 2, 3, ...)
    x <- 1:k
    
    # Calculate column totals (nj)
    nj <- colSums(cont_table)
    
    # For binary dependent variable, use first row as "successes"
    # If more than 2 rows, we need to define what constitutes a "success"
    if (nrow(cont_table) == 2) {
      # Binary case: use first row as successes
      yj <- cont_table[1, ]
    } else {
      # Multi-category case: user must specify which level represents "success"
      # For now, use the first level as reference
      warning("Multi-category dependent variable detected. Using first level as reference category.")
      yj <- cont_table[1, ]
    }
    
    # Calculate observed response probabilities
    pj <- yj / nj
    
    # Calculate weighted mean of scores
    x_bar <- sum(nj * x) / N
    
    # Calculate numerator: Σyj(xj - x̄)
    numerator <- sum(yj * (x - x_bar))
    
    # Calculate denominator components
    p_bar <- sum(yj) / N  # Overall proportion
    
    # Calculate variance term
    variance_term <- p_bar * (1 - p_bar) * sum(nj * (x - x_bar)^2)
    
    # Apply continuity correction if requested
    c <- 0
    if (continuity_correction) {
      if (length(unique(diff(x))) == 1) {
        # Equally spaced scores
        c <- diff(x)[1] / 2
      } else {
        # Not equally spaced - use alternative correction
        c <- min(diff(x)) / 2
        warning("Scores are not equally spaced. Using minimum spacing for continuity correction.")
      }
    }
    
    # Calculate z-statistic with optional continuity correction
    if (test_type == "one.sided.decreasing") {
      # For decreasing trend, add the correction
      z <- (numerator + c) / sqrt(variance_term)
    } else {
      # For two-sided or increasing trend, subtract the correction
      z <- (numerator - c) / sqrt(variance_term)
    }
    
    # Calculate p-value based on test type
    if (test_type == "two.sided") {
      p_value <- 2 * (1 - pnorm(abs(z)))
    } else if (test_type == "one.sided.increasing") {
      p_value <- 1 - pnorm(z)
    } else if (test_type == "one.sided.decreasing") {
      p_value <- pnorm(z)
    }
    
    # Calculate chi-square statistic (z^2)
    chi_square <- z^2
    
    # Create comprehensive result object
    result <- list(
      statistic = c("Z" = z, "Chi-square" = chi_square),
      p.value = p_value,
      method = "Cochran-Armitage Trend Test (Accurate Implementation)",
      data.name = "Contingency table",
      parameters = c(df = 1),
      test_type = test_type,
      continuity_correction = continuity_correction,
      correction_factor = c,
      N = N,
      k = k,
      x_scores = x,
      nj = nj,
      yj = yj,
      pj = pj,
      x_bar = x_bar,
      p_bar = p_bar,
      numerator = numerator,
      variance_term = variance_term
    )
    class(result) <- "htest"
    
    return(result)
  }
  
  # Analysis results
  analysis_results <- eventReactive(input$analyze, {
    req(data_cleaned(), input$dependent_var, input$independent_var)
    
    tryCatch({
      df <- data_cleaned()
      dep_var <- input$dependent_var
      ind_var <- input$independent_var
      
      # Check if selected variables exist
      if (!dep_var %in% names(df) || !ind_var %in% names(df)) {
        stop("Selected variables not found in the dataset")
      }
      
      # Get current level orders
      current_dep_order <- if(!is.null(dep_level_order())) dep_level_order() else unique(df[[dep_var]])
      current_ind_order <- if(!is.null(ind_level_order())) ind_level_order() else unique(df[[ind_var]])
      
      # Remove any remaining missing values specifically for selected variables
      df_analysis <- df %>%
        drop_na(all_of(c(dep_var, ind_var)))
      
      # Check if we have enough data after final cleaning
      if(nrow(df_analysis) == 0) {
        stop("No complete cases available for the selected variables after removing missing values.")
      }
      
      if(nrow(df_analysis) < 5) {
        warning("Small sample size may affect test reliability")
      }
      
      # Convert to factors with custom ordering
      df_processed <- df_analysis %>%
        mutate(
          across(all_of(dep_var), ~factor(., levels = current_dep_order)),
          across(all_of(ind_var), ~factor(., levels = current_ind_order))
        )
      
      # Check factor levels
      if(length(levels(df_processed[[dep_var]])) < 2) {
        stop("Dependent variable must have at least 2 levels.")
      }
      if(length(levels(df_processed[[ind_var]])) < 2) {
        stop("Independent variable must have at least 2 levels.")
      }
      
      # Check if dependent variable is binary
      dep_levels <- levels(df_processed[[dep_var]])
      if(length(dep_levels) > 2) {
        showNotification("Note: Multi-category dependent variable detected. Using first level as reference.", 
                         type = "warning", duration = 5)
      }
      
      # Create contingency table
      cont_table <- table(df_processed[[dep_var]], df_processed[[ind_var]])
      
      # Check table dimensions
      if(any(dim(cont_table) < 2)) {
        stop("Contingency table must have at least 2 rows and 2 columns.")
      }
      
      # Check for sparse data
      if(any(cont_table < 5)) {
        warning("Some cells have counts less than 5. Results should be interpreted with caution.")
      }
      
      # Perform accurate Cochran-Armitage test
      ca_test <- perform_cochran_armitage_accurate(
        cont_table, 
        test_type = input$test_type,
        continuity_correction = input$continuity_correction
      )
      
      # Create stacked bar plot
      bar_plot_data <- df_processed %>%
        count(!!sym(ind_var), !!sym(dep_var)) %>%
        group_by(!!sym(ind_var)) %>%
        mutate(proportion = n / sum(n))
      
      bar_plot <- ggplot(bar_plot_data, aes(x = !!sym(ind_var), y = proportion, fill = !!sym(dep_var))) +
        geom_bar(stat = "identity", position = "stack") +
        geom_text(aes(label = paste0("n=", n)), 
                  position = position_stack(vjust = 0.5), size = 3) +
        labs(
          title = paste("Stacked Bar Plot:", dep_var, "by", ind_var),
          x = ind_var,
          y = "Proportion",
          fill = dep_var
        ) +
        theme_minimal() +
        theme(
          axis.text.x = element_text(angle = 45, hjust = 1),
          legend.position = "bottom"
        ) +
        scale_y_continuous(labels = scales::percent)
      
      # Set analysis as done
      analysis_done(TRUE)
      
      list(
        contingency_table = cont_table,
        test_results = ca_test,
        bar_plot = bar_plot,
        data = df_processed,
        dep_var = dep_var,
        ind_var = ind_var,
        dep_levels = levels(df_processed[[dep_var]]),
        ind_levels = levels(df_processed[[ind_var]]),
        sample_size = nrow(df_processed),
        original_sample_size = nrow(df)
      )
    }, error = function(e) {
      showNotification(paste("Analysis error:", e$message), type = "error")
      return(NULL)
    })
  })
  
  # Test results
  output$test_results <- renderPrint({
    req(analysis_results())
    results <- analysis_results()
    
    cat("COCHRAN-ARMITAGE TEST FOR TREND\n")
    cat("===============================\n\n")
    
    cat("DATA SUMMARY:\n")
    cat("- Original sample size:", results$original_sample_size, "\n")
    cat("- Final sample size (after cleaning):", results$sample_size, "\n")
    cat("- Rows removed due to missing values:", results$original_sample_size - results$sample_size, "\n\n")
    
    cat("VARIABLES:\n")
    cat("- Dependent:", results$dep_var, "\n")
    cat("- Independent:", results$ind_var, "\n")
    cat("- Test type:", results$test_results$test_type, "\n")
    cat("- Continuity correction:", ifelse(results$test_results$continuity_correction, "Yes", "No"), "\n\n")
    
    cat("LEVEL ORDERING:\n")
    cat("- Dependent:", paste(results$dep_levels, collapse = " -> "), "\n")
    cat("- Independent:", paste(results$ind_levels, collapse = " < "), "\n\n")
    
    cat("CONTINGENCY TABLE:\n")
    print(results$contingency_table)
    cat("\n")
    
    cat("TEST CALCULATIONS (Accurate Formula):\n")
    cat("- Total sample size (N):", results$test_results$N, "\n")
    cat("- Number of ordered groups (k):", results$test_results$k, "\n")
    cat("- Column scores (xj):", paste(results$test_results$x_scores, collapse = ", "), "\n")
    cat("- Column totals (nj):", paste(results$test_results$nj, collapse = ", "), "\n")
    cat("- Success counts (yj):", paste(results$test_results$yj, collapse = ", "), "\n")
    cat("- Response probabilities (pj):", paste(format(results$test_results$pj, digits = 4), collapse = ", "), "\n")
    cat("- Weighted mean of scores (x̄):", format(results$test_results$x_bar, digits = 6), "\n")
    cat("- Overall proportion (p̄):", format(results$test_results$p_bar, digits = 6), "\n")
    cat("- Numerator Σyj(xj - x̄):", format(results$test_results$numerator, digits = 6), "\n")
    cat("- Variance term:", format(results$test_results$variance_term, digits = 6), "\n")
    if (results$test_results$continuity_correction) {
      cat("- Continuity correction factor (c):", format(results$test_results$correction_factor, digits = 6), "\n")
    }
    cat("\n")
    
    cat("TEST RESULTS:\n")
    cat("Method:", results$test_results$method, "\n")
    cat("Z-statistic:", format(results$test_results$statistic["Z"], digits = 4), "\n")
    cat("Chi-square statistic:", format(results$test_results$statistic["Chi-square"], digits = 4), "\n")
    cat("Degrees of freedom: 1\n")
    cat("p-value:", format(results$test_results$p.value, digits = 3), "\n\n")
    
    cat("INTERPRETATION:\n")
    p_val <- results$test_results$p.value
    
    if (p_val < 0.05) {
      cat("✓ Statistically significant trend detected (p =", 
          format(p_val, digits = 3), ")\n")
      if (input$test_type == "one.sided.increasing") {
        cat("  Evidence supports an INCREASING trend in", results$dep_var, "across", results$ind_var, "\n")
      } else if (input$test_type == "one.sided.decreasing") {
        cat("  Evidence supports a DECREASING trend in", results$dep_var, "across", results$ind_var, "\n")
      } else {
        cat("  Evidence suggests a systematic relationship between", results$dep_var, "and", results$ind_var, "\n")
      }
    } else {
      cat("✗ No statistically significant trend detected (p =", 
          format(p_val, digits = 3), ")\n")
      cat("  No evidence of systematic relationship between", results$dep_var, 
          "and", results$ind_var, "\n")
    }
    
    # Statistical significance interpretation
    cat("\nSTATISTICAL SIGNIFICANCE:\n")
    if (p_val < 0.001) {
      cat("*** p < 0.001 (Highly significant)\n")
    } else if (p_val < 0.01) {
      cat("** p < 0.01 (Very significant)\n")
    } else if (p_val < 0.05) {
      cat("* p < 0.05 (Significant)\n")
    } else if (p_val < 0.1) {
      cat("~ p < 0.1 (Marginally significant)\n")
    } else {
      cat("p ≥ 0.05 (Not significant)\n")
    }
    
    # Data quality notes
    cat("\nDATA QUALITY NOTES:\n")
    if (results$sample_size < 30) {
      cat("⚠️  Small sample size - results should be interpreted with caution\n")
    }
    if (any(results$contingency_table < 5)) {
      cat("⚠️  Some cell counts < 5 - consider combining categories\n")
    }
    if (length(results$dep_levels) > 2) {
      cat("⚠️  Multi-category dependent variable - using first level as reference\n")
    }
  })
  
  # Bar plot
  output$bar_plot <- renderPlot({
    req(analysis_results())
    results <- analysis_results()
    print(results$bar_plot)
  })
  
  # Report preview
  output$report_preview <- renderUI({
    req(analysis_results())
    
    results <- analysis_results()
    p_val <- results$test_results$p.value
    
    HTML(paste(
      "<h4>Report Preview</h4>",
      "<div style='background-color: #f8f9fa; padding: 15px; border-radius: 5px;'>",
      "<p><strong>Dependent Variable:</strong>", results$dep_var, "</p>",
      "<p><strong>Independent Variable:</strong>", results$ind_var, "</p>",
      "<p><strong>Sample Size:</strong>", results$sample_size, "</p>",
      "<p><strong>Test Type:</strong>", results$test_results$test_type, "</p>",
      "<p><strong>Continuity Correction:</strong>", ifelse(results$test_results$continuity_correction, "Yes", "No"), "</p>",
      "<p><strong>Z-statistic:</strong>", format(results$test_results$statistic["Z"], digits = 4), "</p>",
      "<p><strong>P-value:</strong>", format(p_val, digits = 3), "</p>",
      "<p><strong>Significance:</strong> <span style='color:", 
      if (p_val < 0.05) "#28a745" else "#dc3545", "'>",
      if (p_val < 0.05) "Statistically Significant" else "Not Significant",
      "</span></p>",
      if (results$sample_size < 30) "<p style='color: #856404;'>⚠️ Note: Small sample size</p>" else "",
      "</div>"
    ))
  })
  
  # Download report - FIXED FORMULA VERSION
  output$download_report <- downloadHandler(
    filename = function() {
      paste0("cochran_armitage_report_", Sys.Date(), ".docx")
    },
    
    content = function(file) {
      req(analysis_results())
      results <- analysis_results()
      
      # Show notification that report is being generated
      showNotification("Generating report... This may take a few moments.", 
                       type = "message", duration = 5)
      
      tryCatch({
        # Create a temporary Rmd file
        temp_report <- file.path(tempdir(), "report.Rmd")
        
        # Extract values for the report
        p_val <- results$test_results$p.value
        z_stat <- results$test_results$statistic["Z"]
        chi_sq <- results$test_results$statistic["Chi-square"]
        cont_table <- results$contingency_table
        dep_var <- results$dep_var
        ind_var <- results$ind_var
        sample_size <- results$sample_size
        test_type <- results$test_results$test_type
        continuity_correction <- results$test_results$continuity_correction
        
        # Create the Rmd content with proper formula formatting for Word
        rmd_content <- paste0(
          "---\n",
          "title: \"Cochran-Armitage Trend Test Report\"\n",
          "author: \"CATrend Analyzer\"\n",
          "date: \"", Sys.Date(), "\"\n",
          "output: word_document\n",
          "---\n\n",
          
          "# Cochran-Armitage Trend Test Analysis Report\n\n",
          
          "## Analysis Overview\n\n",
          "- **Analysis Date:** ", Sys.Date(), "\n",
          "- **Dependent Variable:** ", dep_var, "\n",
          "- **Independent Variable:** ", ind_var, "\n", 
          "- **Sample Size:** ", sample_size, "\n",
          "- **Test Type:** ", test_type, "\n",
          "- **Continuity Correction:** ", ifelse(continuity_correction, "Yes", "No"), "\n\n",
          
          "## Level Ordering\n\n",
          "- **Dependent Variable Levels:** ", paste(results$dep_levels, collapse = " → "), "\n",
          "- **Independent Variable Levels:** ", paste(results$ind_levels, collapse = " < "), "\n\n",
          
          "## Contingency Table\n\n",
          "```{r, echo=FALSE}\n",
          "knitr::kable(addmargins(cont_table), caption = 'Contingency Table with Totals')\n",
          "```\n\n",
          
          "## Statistical Results\n\n",
          "- **Z-statistic:** ", format(z_stat, digits = 4), "\n",
          "- **Chi-square statistic:** ", format(chi_sq, digits = 4), "\n", 
          "- **P-value:** ", format(p_val, digits = 4), "\n\n",
          
          "## Interpretation\n\n",
          if (p_val < 0.05) {
            paste0("There is a statistically significant trend in the proportion of **", 
                   dep_var, "** across the ordered levels of **", 
                   ind_var, "** (Z = ", format(z_stat, digits = 4), 
                   ", p = ", format(p_val, digits = 4), "). ",
                   "This suggests a systematic relationship between the variables.")
          } else {
            paste0("No statistically significant trend was detected in the proportion of **",
                   dep_var, "** across the ordered levels of **",
                   ind_var, "** (Z = ", format(z_stat, digits = 4),
                   ", p = ", format(p_val, digits = 4), "). ",
                   "This suggests no systematic relationship between the variables.")
          },
          "\n\n",
          
          "## Methodological Details\n\n",
          "The Cochran-Armitage test for trend was performed using the following formula:\n\n",
          "Z = [Σ y_j (x_j - x̄)] / √[p̄(1-p̄) Σ n_j (x_j - x̄)²]\n\n",
          "Where:\n",
          "- x_j are the scores for the ordered groups\n", 
          "- y_j are the success counts in each group\n",
          "- n_j are the total counts in each group\n",
          "- x̄ is the weighted mean of scores\n",
          "- p̄ is the overall proportion of successes\n\n",
          
          "---\n",
          "*Report generated using CATrend Analyzer*"
        )
        
        # Write the Rmd content to temporary file
        writeLines(rmd_content, temp_report)
        
        # Save the plot to a temporary file
        temp_plot <- file.path(tempdir(), "plot.png")
        ggsave(temp_plot, plot = results$bar_plot, width = 8, height = 6, dpi = 300)
        
        # Render the report
        render(temp_report, 
               output_file = file,
               envir = new.env(),
               quiet = TRUE)
        
        showNotification("Report generated successfully!", type = "message")
        
      }, error = function(e) {
        showNotification(paste("Error generating report:", e$message), type = "error")
      })
    }
  )
}

# Run the application
shinyApp(ui = ui, server = server)
