# app.R
library(shiny)
library(shinythemes)
library(readr)
library(readxl)
library(haven)
library(DT)
library(flextable)
library(officer)
library(dplyr)
library(purrr)
library(stringr)
library(papaja)
library(sjPlot)
library(summarytools)
library(ggplot2)

ui <- fluidPage(
  theme = shinytheme("cosmo"),
  
  # Custom CSS for enhanced styling
  tags$head(
    tags$style(HTML("
      .navbar { background-color: #2C3E50; }
      .navbar-default .navbar-nav > .active > a { 
        background-color: #34495E !important; 
        color: white !important;
      }
      .well {
        background-color: #ECF0F1;
        border: 1px solid #BDC3C7;
      }
      .btn-primary {
        background-color: #3498DB;
        border-color: #2980B9;
      }
      .btn-success {
        background-color: #27AE60;
        border-color: #229954;
      }
      .btn-load {
        background-color: #E67E22;
        border-color: #D35400;
        color: white;
        width: 100%;
        padding: 10px;
        font-size: 16px;
        font-weight: bold;
      }
      .table-apa {
        font-family: 'Times New Roman', serif;
        font-size: 12pt;
      }
      .shiny-notification {
        position: fixed;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
      }
      .data-info {
        background-color: #F8F9FA;
        padding: 15px;
        border-radius: 5px;
        margin-top: 10px;
        border-left: 4px solid #3498DB;
      }
      .file-input-info {
        background-color: #FFF3CD;
        padding: 10px;
        border-radius: 5px;
        border-left: 4px solid #FFC107;
        margin-bottom: 10px;
      }
      .selected-vars {
        background-color: #D5EDDA;
        padding: 10px;
        border-radius: 5px;
        margin-top: 10px;
      }
      .stats-options-group {
        max-height: 300px;
        overflow-y: auto;
        border: 1px solid #ddd;
        padding: 10px;
        border-radius: 5px;
        background-color: #f9f9f9;
      }
      .home-panel {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 30px;
        border-radius: 10px;
        margin-bottom: 20px;
      }
      .feature-card {
        background-color: white;
        border-radius: 8px;
        padding: 20px;
        margin-bottom: 15px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        border-left: 4px solid #3498DB;
      }
      .copyright-footer {
        position: fixed;
        bottom: 0;
        left: 0;
        right: 0;
        background-color: #2C3E50;
        color: white;
        padding: 10px;
        text-align: center;
        font-size: 12px;
        z-index: 1000;
      }
      .main-content {
        margin-bottom: 60px; /* Space for footer */
      }
      .crosstab-header {
        background-color: #34495E;
        color: white;
        font-weight: bold;
        text-align: center;
      }
      .warning-panel {
        background-color: #FFF3CD;
        border: 1px solid #FFC107;
        border-radius: 5px;
        padding: 15px;
        margin-bottom: 15px;
      }
      .stats-options-container {
        max-height: 400px;
        overflow-y: auto;
        padding-right: 10px;
      }
      .scrollable-well {
        max-height: 500px;
        overflow-y: auto;
        padding-right: 10px;
      }
    "))
  ),
  
  # Copyright Footer
  tags$div(
    class = "copyright-footer",
    "© 2025 Mudasir Mohammed Ibrahim (mudassiribrahim30@gmail.com) - APA Table Generator Pro"
  ),
  
  # Main Content
  tags$div(
    class = "main-content",
    
    # Navigation Bar
    navbarPage(
      "APA Table Generator Pro",
      id = "main_nav",
      theme = shinytheme("cosmo"),
      
      # Home Tab
      tabPanel(
        "Home",
        icon = icon("home"),
        fluidRow(
          column(
            12,
            div(
              class = "home-panel",
              h1("APA Table Generator Pro", style = "color: white; font-weight: bold;"),
              h4("Professional Statistical Tables in APA Format", style = "color: white; opacity: 0.9;"),
              p("Create publication-ready APA formatted tables with ease. Import your data, select variables, and generate perfectly formatted tables for your research papers and reports.", 
                style = "color: white; font-size: 16px; margin-top: 20px;")
            )
          )
        ),
        
        fluidRow(
          column(
            4,
            div(
              class = "feature-card",
              h4(icon("upload"), "Easy Data Import"),
              p("Support for multiple file formats: CSV, Excel, SPSS, Stata, SAS, and RData. Automatic data type detection and cleaning.")
            )
          ),
          column(
            4,
            div(
              class = "feature-card",
              h4(icon("sliders-h"), "Flexible Analysis"),
              p("Choose from descriptive statistics, frequency tables, or cross-tabulation. Customize your analysis with various statistical options.")
            )
          ),
          column(
            4,
            div(
              class = "feature-card",
              h4(icon("file-alt"), "APA Formatting"),
              p("Generate publication-ready tables in APA 7th edition format. Multiple table styles and automatic captioning.")
            )
          )
        ),
        
        fluidRow(
          column(
            6,
            div(
              class = "feature-card",
              h4(icon("sort"), "Smart Sorting"),
              p("Sort variable levels by frequency in ascending or descending order. Perfect for organizing categorical data in your tables.")
            )
          ),
          column(
            6,
            div(
              class = "feature-card",
              h4(icon("download"), "Export Options"),
              p("Download your tables as Word documents. Perfect for direct inclusion in research papers and reports.")
            )
          )
        ),
        
        fluidRow(
          column(
            12,
            wellPanel(
              h4("How to Use This App", icon("info-circle")),
              tags$ol(
                tags$li(strong("Data Import:"), "Go to the Data Import tab, upload your data file, and click 'Load Data'"),
                tags$li(strong("Variable Selection:"), "Select your variables and choose the type of analysis (Descriptive, Frequency, or Cross-tabulation)"),
                tags$li(strong("Statistics Options:"), "Customize which statistics to display based on your analysis type"),
                tags$li(strong("Sorting:"), "Optionally sort variable levels by frequency in ascending or descending order"),
                tags$li(strong("Generate Table:"), "Go to APA Table Results to view and download your formatted table")
              ),
              br(),
              h5("Supported File Formats:"),
              p("CSV, Excel (.xlsx, .xls), SPSS (.sav), Stata (.dta), SAS (.sas7bdat), RData (.rda, .rdata)"),
              br(),
              h5("Need Help?"),
              p("For technical support or feature requests, contact: mudassiribrahim30@gmail.com")
            )
          )
        )
      ),
      
      # Data Import Tab
      tabPanel(
        "Data Import",
        icon = icon("upload"),
        fluidRow(
          column(
            12,
            wellPanel(
              h4("Upload Your Data File", icon("file-import")),
              
              # File upload section
              fluidRow(
                column(8,
                       fileInput("file_upload", NULL,
                                 multiple = FALSE,
                                 accept = c(
                                   ".csv", ".xlsx, .xls", 
                                   ".dta", ".sav", ".sas7bdat", ".rdata", ".rda"
                                 ),
                                 buttonLabel = "Browse Files",
                                 placeholder = "No file selected")
                ),
                column(4,
                       # Always visible Load Data button
                       actionButton("load_data", "Load Data", 
                                    class = "btn-load",
                                    icon = icon("database"))
                )
              ),
              
              # File info and options
              uiOutput("file_info_panel"),
              
              # Data format options
              conditionalPanel(
                condition = "input.file_upload != null",
                fluidRow(
                  column(4,
                         selectInput("file_type", "File Type:",
                                     choices = c("Auto-detect" = "auto",
                                                 "CSV" = "csv",
                                                 "Excel" = "excel",
                                                 "Stata" = "stata",
                                                 "SPSS" = "spss",
                                                 "SAS" = "sas",
                                                 "RData" = "rdata"),
                                     selected = "auto")
                  ),
                  column(4,
                         conditionalPanel(
                           condition = "input.file_type == 'csv'",
                           checkboxInput("header", "Header", TRUE),
                           selectInput("sep", "Separator",
                                       choices = c(Comma = ",", Semicolon = ";", Tab = "\t"),
                                       selected = ","),
                           selectInput("encoding", "Encoding",
                                       choices = c("UTF-8" = "UTF-8", 
                                                   "Latin-1" = "Latin-1",
                                                   "ASCII" = "ASCII"),
                                       selected = "UTF-8")
                         )
                  ),
                  column(4,
                         uiOutput("data_status")
                  )
                )
              )
            )
          )
        ),
        
        fluidRow(
          column(
            12,
            wellPanel(
              h4("Data Preview", icon("table")),
              DTOutput("data_preview")
            )
          )
        )
      ),
      
      # Variable Selection Tab
      tabPanel(
        "Variable Selection",
        icon = icon("sliders-h"),
        fluidRow(
          column(
            6,
            wellPanel(
              h4("Select Variables", icon("check-circle")),
              uiOutput("variable_selector"),
              br(),
              uiOutput("selected_vars_display"),
              fluidRow(
                column(6,
                       actionButton("clear_selection", "Clear Selection", 
                                    class = "btn-warning btn-block", 
                                    icon = icon("times"))
                ),
                column(6,
                       actionButton("select_numeric", "Select Numeric Only", 
                                    class = "btn-info btn-block", 
                                    icon = icon("sort-numeric-up"))
                )
              )
            )
          ),
          column(
            6,
            div(
              class = "scrollable-well",
              wellPanel(
                h4("Statistics Options", icon("calculator")),
                radioButtons("table_type", "Table Type:",
                             choices = c("Descriptive Statistics" = "descriptive",
                                         "Frequency Table" = "frequency",
                                         "Cross Tabulation" = "crosstab"),
                             selected = "descriptive"),
                
                # Sorting options
                conditionalPanel(
                  condition = "input.table_type == 'frequency' || input.table_type == 'crosstab'",
                  h5("Sorting Options:"),
                  radioButtons("sort_order", "Sort Variable Levels:",
                               choices = c("None" = "none",
                                           "Ascending by Frequency" = "asc",
                                           "Descending by Frequency" = "desc"),
                               selected = "none")
                ),
                
                # Conditional statistics options based on table type
                conditionalPanel(
                  condition = "input.table_type == 'descriptive'",
                  h5("Descriptive Statistics Options:"),
                  div(class = "stats-options-group",
                      checkboxGroupInput("desc_stats_options", "Select Statistics:",
                                         choices = c("Mean" = "mean",
                                                     "Standard Deviation" = "sd",
                                                     "Median" = "median",
                                                     "Mode" = "mode",
                                                     "Minimum" = "min",
                                                     "Maximum" = "max",
                                                     "Range" = "range",
                                                     "Skewness" = "skewness",
                                                     "Kurtosis" = "kurtosis",
                                                     "Variance" = "variance",
                                                     "Standard Error" = "se",
                                                     "Count (N)" = "count"),
                                         selected = c("mean", "sd", "min", "max"))
                  )
                ),
                
                conditionalPanel(
                  condition = "input.table_type == 'frequency'",
                  h5("Frequency Table Options:"),
                  div(class = "stats-options-group",
                      checkboxGroupInput("freq_stats_options", "Select Statistics:",
                                         choices = c("Frequency" = "freq",
                                                     "Percentage" = "percent",
                                                     "Cumulative Frequency" = "cum_freq",
                                                     "Cumulative Percentage" = "cum_percent",
                                                     "Valid Percentage" = "valid_percent",
                                                     "Include Total Row" = "total"),
                                         selected = c("freq", "percent"))
                  ),
                  
                  # Continuous variables handling for frequency tables
                  uiOutput("continuous_vars_warning"),
                  conditionalPanel(
                    condition = "output.has_continuous_vars == true && input.table_type == 'frequency'",
                    div(
                      class = "warning-panel",
                      h5("Continuous Variables Detected", icon("exclamation-triangle")),
                      p("You have selected continuous variables for a frequency table. Would you like to include descriptive statistics for these variables?"),
                      radioButtons("include_continuous", "Include Continuous Variables:",
                                   choices = c("No, show only categorical variables" = "no",
                                               "Yes, include descriptive statistics" = "yes"),
                                   selected = "no"),
                      conditionalPanel(
                        condition = "input.include_continuous == 'yes'",
                        h6("Select Descriptive Statistics for Continuous Variables:"),
                        checkboxGroupInput("continuous_stats", "Statistics to Display:",
                                           choices = c("Mean" = "mean",
                                                       "Standard Deviation" = "sd",
                                                       "Median" = "median",
                                                       "Minimum" = "min",
                                                       "Maximum" = "max",
                                                       "Range" = "range",
                                                       "Count (N)" = "count"),
                                           selected = c("mean", "sd"))
                      )
                    )
                  )
                ),
                
                conditionalPanel(
                  condition = "input.table_type == 'crosstab'",
                  h5("Cross Tabulation Options:"),
                  p("Note: For cross tabulation, select exactly 2 variables", 
                    style = "color: #E74C3C; font-style: italic;"),
                  div(class = "stats-options-group",
                      checkboxGroupInput("crosstab_stats_options", "Select Statistics:",
                                         choices = c("Count" = "count",
                                                     "Row Percentage" = "row_percent",
                                                     "Column Percentage" = "col_percent",
                                                     "Total Percentage" = "total_percent",
                                                     "Chi-square Test" = "chisq",
                                                     "Include Totals" = "include_totals"),
                                         selected = c("count", "row_percent", "col_percent", "total_percent"))
                  )
                )
              )
            )
          )
        ),
        
        fluidRow(
          column(
            12,
            wellPanel(
              h4("Variable Information", icon("info-circle")),
              verbatimTextOutput("var_info")
            )
          )
        )
      ),
      
      # APA Table Results Tab
      tabPanel(
        "APA Table Results",
        icon = icon("file-alt"),
        fluidRow(
          column(
            12,
            wellPanel(
              fluidRow(
                column(6,
                       textInput("table_caption", "Table Caption:",
                                 placeholder = "e.g., Table 1: Descriptive Statistics of Study Variables")
                ),
                column(3,
                       numericInput("table_number", "Table Number:", 
                                    value = 1, min = 1, max = 100)
                ),
                column(3,
                       selectInput("table_style", "Table Style:",
                                   choices = c("APA 7th" = "apa",
                                               "Modern" = "modern",
                                               "Classic" = "classic"),
                                   selected = "apa")
                )
              )
            )
          )
        ),
        
        fluidRow(
          column(
            12,
            wellPanel(
              h4("APA Formatted Table", icon("table")),
              div(class = "table-apa",
                  uiOutput("apa_table_output")),
              br(),
              downloadButton("download_doc", "Download Word Document",
                             class = "btn-success btn-block",
                             icon = icon("file-word"))
            )
          )
        )
      )
    )
  )
)

server <- function(input, output, session) {
  
  # Reactive value for stored data
  data_reactive <- reactiveVal(NULL)
  selected_vars_reactive <- reactiveVal(character(0))
  
  # Reactive value to track variable types
  var_types_reactive <- reactive({
    req(data_reactive(), selected_vars_reactive())
    data <- data_reactive()
    selected_vars <- selected_vars_reactive()
    
    var_types <- sapply(selected_vars, function(var) {
      if(is.numeric(data[[var]])) "continuous" else "categorical"
    })
    
    return(var_types)
  })
  
  # Check if there are continuous variables selected for frequency table
  output$has_continuous_vars <- reactive({
    req(var_types_reactive())
    any(var_types_reactive() == "continuous") && input$table_type == "frequency"
  })
  outputOptions(output, "has_continuous_vars", suspendWhenHidden = FALSE)
  
  # Continuous variables warning and options
  output$continuous_vars_warning <- renderUI({
    req(var_types_reactive())
    
    continuous_vars <- names(var_types_reactive()[var_types_reactive() == "continuous"])
    categorical_vars <- names(var_types_reactive()[var_types_reactive() == "categorical"])
    
    if(length(continuous_vars) > 0 && input$table_type == "frequency") {
      tagList(
        p(strong("Continuous variables detected:"), paste(continuous_vars, collapse = ", ")),
        p(strong("Categorical variables:"), if(length(categorical_vars) > 0) paste(categorical_vars, collapse = ", ") else "None")
      )
    }
  })
  
  # File info panel
  output$file_info_panel <- renderUI({
    if(is.null(input$file_upload)) {
      tags$div(
        class = "file-input-info",
        icon("info-circle"),
        "Please select a data file and click 'Load Data' to import your data."
      )
    } else {
      tags$div(
        class = "file-input-info",
        icon("file"),
        strong("File selected:"), input$file_upload$name,
        br(),
        strong("Size:"), format(input$file_upload$size, big.mark = ","), "bytes",
        br(),
        strong("Type:"), tools::file_ext(input$file_upload$name)
      )
    }
  })
  
  # Data status display
  output$data_status <- renderUI({
    if(is.null(data_reactive())) {
      if(!is.null(input$file_upload)) {
        tags$div(
          class = "alert alert-warning",
          icon("exclamation-triangle"),
          "Data not loaded yet. Click 'Load Data'."
        )
      } else {
        tags$div(
          class = "alert alert-info",
          icon("info-circle"),
          "No file selected."
        )
      }
    } else {
      data <- data_reactive()
      tags$div(
        class = "data-info",
        h5("Data Loaded Successfully!", style = "color: #27AE60;"),
        p(icon("table"), strong("Variables:"), ncol(data)),
        p(icon("bars"), strong("Observations:"), nrow(data)),
        p(icon("check-circle"), strong("Status:"), "Ready for analysis")
      )
    }
  })
  
  # Read data when load button is clicked
  observeEvent(input$load_data, {
    # Check if file is selected
    if(is.null(input$file_upload)) {
      showNotification("Please select a file first!", type = "warning", duration = 5)
      return()
    }
    
    file_path <- input$file_upload$datapath
    file_ext <- tolower(tools::file_ext(input$file_upload$name))
    file_type <- input$file_type
    
    showNotification("Loading data... Please wait.", type = "message", duration = NULL, id = "loading")
    
    tryCatch({
      data_loaded <- NULL
      
      # Use specified file type or auto-detect
      if(file_type != "auto") {
        file_ext <- file_type
      }
      
      data_loaded <- switch(
        file_ext,
        "csv" = {
          read_csv(file_path, 
                   col_names = input$header,
                   locale = locale(encoding = input$encoding))
        },
        "excel" = read_excel(file_path),
        "stata" = read_dta(file_path),
        "spss" = read_sav(file_path),
        "sas" = read_sas(file_path),
        "rdata" = {
          env <- new.env()
          load(file_path, envir = env)
          obj_names <- ls(env)
          if(length(obj_names) > 0) {
            get(obj_names[1], envir = env)
          } else {
            stop("No objects found in RData file")
          }
        },
        # Auto-detect based on file extension
        {
          switch(file_ext,
                 "xlsx" = read_excel(file_path),
                 "xls" = read_excel(file_path),
                 "dta" = read_dta(file_path),
                 "sav" = read_sav(file_path),
                 "sas7bdat" = read_sas(file_path),
                 "rda" = {
                   env <- new.env()
                   load(file_path, envir = env)
                   obj_names <- ls(env)
                   if(length(obj_names) > 0) {
                     get(obj_names[1], envir = env)
                   } else {
                     stop("No objects found in RData file")
                   }
                 },
                 # Default to CSV
                 read_csv(file_path, col_names = TRUE)
          )
        }
      )
      
      # Convert to data frame if it's not already
      if(!is.data.frame(data_loaded)) {
        data_loaded <- as.data.frame(data_loaded)
      }
      
      # Remove empty rows and columns
      data_loaded <- data_loaded[, colSums(!is.na(data_loaded)) > 0]
      data_loaded <- data_loaded[rowSums(!is.na(data_loaded)) > 0, ]
      
      data_reactive(data_loaded)
      selected_vars_reactive(character(0)) # Reset selection
      
      removeNotification("loading")
      showNotification(
        paste("Data loaded successfully! Found", ncol(data_loaded), "variables and", nrow(data_loaded), "observations."),
        type = "message", 
        duration = 5
      )
      
      # Update to Variable Selection tab
      updateNavbarPage(session, "main_nav", selected = "Variable Selection")
      
    }, error = function(e) {
      removeNotification("loading")
      showNotification(
        paste("Error loading file:", e$message), 
        type = "error", 
        duration = 10
      )
      data_reactive(NULL)
    })
  })
  
  # Data preview
  output$data_preview <- renderDT({
    req(data_reactive())
    data <- data_reactive()
    datatable(
      data,
      options = list(
        scrollX = TRUE, 
        pageLength = 5,
        dom = 'Bfrtip',
        buttons = c('copy', 'csv', 'excel')
      ),
      rownames = FALSE,
      selection = 'none'
    )
  })
  
  # Variable selector UI - Dropdown with multi-select
  output$variable_selector <- renderUI({
    req(data_reactive())
    data <- data_reactive()
    var_names <- names(data)
    
    # Create variable descriptions with types
    var_choices <- var_names
    var_descriptions <- sapply(var_names, function(var) {
      var_type <- ifelse(is.numeric(data[[var]]), " (numeric)", " (categorical)")
      paste0(var, var_type)
    })
    
    selectInput(
      "selected_vars", 
      "Select variables for analysis:",
      choices = setNames(var_names, var_descriptions),
      selected = selected_vars_reactive(),
      multiple = TRUE,
      selectize = TRUE
    )
  })
  
  # Display selected variables
  output$selected_vars_display <- renderUI({
    selected_vars <- selected_vars_reactive()
    if(length(selected_vars) > 0) {
      tags$div(
        class = "selected-vars",
        h5("Selected Variables:"),
        tags$ul(
          lapply(selected_vars, function(var) tags$li(var))
        )
      )
    }
  })
  
  # Update selected variables reactively
  observeEvent(input$selected_vars, {
    selected_vars_reactive(input$selected_vars)
  })
  
  # Clear selection
  observeEvent(input$clear_selection, {
    selected_vars_reactive(character(0))
    updateSelectInput(session, "selected_vars", selected = character(0))
  })
  
  # Select numeric variables only
  observeEvent(input$select_numeric, {
    req(data_reactive())
    data <- data_reactive()
    numeric_vars <- names(data)[sapply(data, is.numeric)]
    selected_vars_reactive(numeric_vars)
    updateSelectInput(session, "selected_vars", selected = numeric_vars)
  })
  
  # Variable information
  output$var_info <- renderPrint({
    req(data_reactive())
    
    data <- data_reactive()
    selected_vars <- selected_vars_reactive()
    
    if(length(selected_vars) == 0) {
      cat("No variables selected. Please select variables from the dropdown above.\n")
      return()
    }
    
    cat("Selected Variables:", length(selected_vars), "\n")
    cat("Table Type:", switch(input$table_type,
                              "descriptive" = "Descriptive Statistics",
                              "frequency" = "Frequency Table", 
                              "crosstab" = "Cross Tabulation"), "\n\n")
    
    for(var in selected_vars) {
      cat("Variable:", var, "\n")
      cat("Type:", ifelse(is.numeric(data[[var]]), "Numeric", "Categorical"), "\n")
      cat("Missing Values:", sum(is.na(data[[var]])), 
          paste0("(", round(mean(is.na(data[[var]])) * 100, 1), "%)"), "\n")
      
      if(is.numeric(data[[var]])) {
        cat("Summary Statistics:\n")
        cat("  Min:", round(min(data[[var]], na.rm = TRUE), 2), "\n")
        cat("  Max:", round(max(data[[var]], na.rm = TRUE), 2), "\n")
        cat("  Mean:", round(mean(data[[var]], na.rm = TRUE), 2), "\n")
        cat("  SD:", round(sd(data[[var]], na.rm = TRUE), 2), "\n")
        cat("  Median:", round(median(data[[var]], na.rm = TRUE), 2), "\n")
      } else {
        cat("Categories:", length(unique(data[[var]])), "\n")
        levels <- head(unique(as.character(data[[var]])), 5)
        cat("First few levels:", paste(levels, collapse = ", "))
        if(length(unique(data[[var]])) > 5) cat("...")
        cat("\n")
      }
      cat("---\n")
    }
  })
  
  # Generate APA table reactively
  apa_table_reactive <- reactive({
    req(data_reactive(), selected_vars_reactive())
    
    data <- data_reactive()
    selected_vars <- selected_vars_reactive()
    table_type <- input$table_type
    sort_order <- input$sort_order
    
    if(length(selected_vars) == 0) {
      return(NULL)
    }
    
    tryCatch({
      if(table_type == "descriptive") {
        stats_options <- input$desc_stats_options
        generate_descriptive_table(data, selected_vars, stats_options)
      } else if(table_type == "frequency") {
        stats_options <- input$freq_stats_options
        include_continuous <- input$include_continuous
        continuous_stats <- input$continuous_stats
        generate_frequency_table(data, selected_vars, stats_options, sort_order, include_continuous, continuous_stats)
      } else if(table_type == "crosstab") {
        stats_options <- input$crosstab_stats_options
        generate_crosstab_table(data, selected_vars, stats_options, sort_order)
      }
    }, error = function(e) {
      showNotification(paste("Error generating table:", e$message), type = "error")
      return(NULL)
    })
  })
  
  # Helper function to calculate mode
  calculate_mode <- function(x) {
    if(length(x) == 0) return(NA)
    if(all(is.na(x))) return(NA)
    ux <- unique(na.omit(x))
    if(length(ux) == 0) return(NA)
    freq <- tabulate(match(na.omit(x), ux))
    modes <- ux[freq == max(freq)]
    if(length(modes) > 1) {
      return(paste(modes[1], "(multiple modes)"))
    }
    return(modes[1])
  }
  
  # Helper function to calculate skewness
  calculate_skewness <- function(x) {
    if(length(na.omit(x)) < 3) return(NA)
    x <- na.omit(x)
    n <- length(x)
    if(sd(x) == 0) return(0)
    (sum((x - mean(x))^3) / n) / (sd(x)^3)
  }
  
  # Helper function to calculate kurtosis
  calculate_kurtosis <- function(x) {
    if(length(na.omit(x)) < 4) return(NA)
    x <- na.omit(x)
    n <- length(x)
    if(sd(x) == 0) return(0)
    (sum((x - mean(x))^4) / n) / (sd(x)^4) - 3
  }
  
  # Generate descriptive statistics table with individual columns
  generate_descriptive_table <- function(data, variables, stats_options) {
    desc_list <- list()
    
    for(var in variables) {
      if(is.numeric(data[[var]])) {
        # Continuous variables
        var_data <- na.omit(data[[var]])
        n_valid <- length(var_data)
        
        if(n_valid == 0) {
          # Handle case with no valid data
          row_data <- list(Variable = var)
          for(stat in stats_options) {
            row_data[[stat]] <- NA
          }
          desc_list[[var]] <- as.data.frame(row_data)
          next
        }
        
        # Calculate all requested statistics
        stats_values <- list()
        
        if("mean" %in% stats_options) {
          stats_values$mean <- mean(var_data)
        }
        if("sd" %in% stats_options) {
          stats_values$sd <- sd(var_data)
        }
        if("median" %in% stats_options) {
          stats_values$median <- median(var_data)
        }
        if("mode" %in% stats_options) {
          stats_values$mode <- calculate_mode(var_data)
        }
        if("min" %in% stats_options) {
          stats_values$min <- min(var_data)
        }
        if("max" %in% stats_options) {
          stats_values$max <- max(var_data)
        }
        if("range" %in% stats_options) {
          stats_values$range <- max(var_data) - min(var_data)
        }
        if("skewness" %in% stats_options) {
          stats_values$skewness <- calculate_skewness(var_data)
        }
        if("kurtosis" %in% stats_options) {
          stats_values$kurtosis <- calculate_kurtosis(var_data)
        }
        if("variance" %in% stats_options) {
          stats_values$variance <- var(var_data)
        }
        if("se" %in% stats_options) {
          stats_values$se <- sd(var_data) / sqrt(length(var_data))
        }
        if("count" %in% stats_options) {
          stats_values$count <- n_valid
        }
        
        # Create row with Variable name and all selected statistics
        row_data <- c(list(Variable = var), stats_values)
        desc_list[[var]] <- as.data.frame(row_data)
        
      } else {
        # Categorical variables - basic descriptive
        var_data <- na.omit(data[[var]])
        n_valid <- length(var_data)
        n_categories <- length(unique(var_data))
        
        row_data <- list(Variable = var)
        if("count" %in% stats_options) {
          row_data$count <- n_valid
        }
        if("mode" %in% stats_options) {
          row_data$mode <- calculate_mode(var_data)
        }
        
        desc_list[[var]] <- as.data.frame(row_data)
      }
    }
    
    if(length(desc_list) == 0) {
      return(data.frame(Message = "No valid data for selected variables"))
    }
    
    result <- bind_rows(desc_list)
    
    # Format column names for better display
    col_names <- names(result)
    col_names <- sapply(col_names, function(name) {
      switch(name,
             "mean" = "M",
             "sd" = "SD",
             "se" = "SE",
             "count" = "N",
             "min" = "Min",
             "max" = "Max",
             "range" = "Range",
             "mode" = "Mode",
             "median" = "Median",
             "variance" = "Variance",
             "skewness" = "Skewness",
             "kurtosis" = "Kurtosis",
             name)
    })
    names(result) <- col_names
    
    return(result)
  }
  
  # Generate frequency table in APA format with sorting and continuous variables support
  generate_frequency_table <- function(data, variables, stats_options, sort_order = "none", 
                                       include_continuous = "no", continuous_stats = NULL) {
    freq_list <- list()
    
    # Process variables in the order they were selected
    for(var in variables) {
      if(!is.numeric(data[[var]])) {
        # Categorical variable processing
        var_data <- data[[var]]
        non_missing_data <- var_data[!is.na(var_data)]
        
        if(length(non_missing_data) == 0) {
          # Handle case with no valid data
          freq_list[[var]] <- data.frame(
            Variable = var,
            Level = "No valid data",
            Frequency = 0,
            Percentage = 0
          )
          next
        }
        
        freq_table <- table(non_missing_data)
        total_count <- sum(freq_table)
        prop_table <- prop.table(freq_table) * 100
        
        # Apply sorting if requested
        if(sort_order != "none") {
          if(sort_order == "asc") {
            # Sort by frequency ascending
            freq_table <- freq_table[order(freq_table)]
            prop_table <- prop_table[order(freq_table)]
          } else if(sort_order == "desc") {
            # Sort by frequency descending
            freq_table <- freq_table[order(freq_table, decreasing = TRUE)]
            prop_table <- prop_table[order(freq_table, decreasing = TRUE)]
          }
        }
        
        # Calculate cumulative statistics if requested
        cum_freq <- cumsum(freq_table)
        cum_percent <- cumsum(prop_table)
        
        for(i in seq_along(freq_table)) {
          level_name <- names(freq_table)[i]
          
          row_data <- list(
            Variable = ifelse(i == 1, var, ""),
            Level = level_name
          )
          
          if("freq" %in% stats_options) {
            row_data$Frequency <- freq_table[i]
          }
          if("percent" %in% stats_options) {
            row_data$Percentage <- round(prop_table[i], 1)
          }
          if("cum_freq" %in% stats_options) {
            row_data$`Cumulative Frequency` <- cum_freq[i]
          }
          if("cum_percent" %in% stats_options) {
            row_data$`Cumulative Percentage` <- round(cum_percent[i], 1)
          }
          if("valid_percent" %in% stats_options) {
            # Same as percentage for non-missing data
            row_data$`Valid Percentage` <- round(prop_table[i], 1)
          }
          
          freq_list[[paste(var, level_name)]] <- as.data.frame(row_data)
        }
        
        # Add total row if requested
        if("total" %in% stats_options) {
          total_row <- list(Variable = "", Level = "Total")
          if("freq" %in% stats_options) total_row$Frequency <- total_count
          if("percent" %in% stats_options) total_row$Percentage <- 100
          if("cum_freq" %in% stats_options) total_row$`Cumulative Frequency` <- total_count
          if("cum_percent" %in% stats_options) total_row$`Cumulative Percentage` <- 100
          if("valid_percent" %in% stats_options) total_row$`Valid Percentage` <- 100
          
          freq_list[[paste(var, "Total")]] <- as.data.frame(total_row)
        }
      } else if(include_continuous == "yes" && !is.null(continuous_stats)) {
        # Continuous variable processing (only if user wants to include them)
        var_data <- na.omit(data[[var]])
        n_valid <- length(var_data)
        
        if(n_valid == 0) {
          # Handle case with no valid data
          row_data <- list(Variable = var, Level = "No valid data")
          for(stat in continuous_stats) {
            row_data[[stat]] <- NA
          }
          freq_list[[paste(var, "continuous")]] <- as.data.frame(row_data)
          next
        }
        
        # Calculate requested statistics
        stats_values <- list()
        
        if("mean" %in% continuous_stats) {
          stats_values$Mean <- round(mean(var_data), 2)
        }
        if("sd" %in% continuous_stats) {
          stats_values$SD <- round(sd(var_data), 2)
        }
        if("median" %in% continuous_stats) {
          stats_values$Median <- round(median(var_data), 2)
        }
        if("min" %in% continuous_stats) {
          stats_values$Minimum <- round(min(var_data), 2)
        }
        if("max" %in% continuous_stats) {
          stats_values$Maximum <- round(max(var_data), 2)
        }
        if("range" %in% continuous_stats) {
          stats_values$Range <- round(max(var_data) - min(var_data), 2)
        }
        if("count" %in% continuous_stats) {
          stats_values$N <- n_valid
        }
        
        # Create row for continuous variable
        row_data <- c(list(Variable = var, Level = "Continuous"), stats_values)
        freq_list[[paste(var, "continuous")]] <- as.data.frame(row_data)
      }
    }
    
    if(length(freq_list) == 0) {
      return(data.frame(Message = "No variables selected for frequency table"))
    }
    
    result <- bind_rows(freq_list)
    return(result)
  }
  
  # Generate cross tabulation table with combined statistics format
  generate_crosstab_table <- function(data, variables, stats_options, sort_order = "none") {
    if(length(variables) < 2) {
      return(data.frame(Message = "Select at least 2 variables for cross tabulation"))
    }
    
    var1 <- variables[1]
    var2 <- variables[2]
    
    # Remove rows with missing values in either variable
    complete_cases <- complete.cases(data[[var1]], data[[var2]])
    var1_data <- data[[var1]][complete_cases]
    var2_data <- data[[var2]][complete_cases]
    
    if(length(var1_data) == 0 || length(var2_data) == 0) {
      return(data.frame(Message = "No complete cases for cross tabulation"))
    }
    
    # Apply sorting if requested
    if(sort_order != "none") {
      # Sort var1 levels by frequency
      var1_freq <- table(var1_data)
      if(sort_order == "asc") {
        var1_levels <- names(var1_freq)[order(var1_freq)]
      } else {
        var1_levels <- names(var1_freq)[order(var1_freq, decreasing = TRUE)]
      }
      var1_data <- factor(var1_data, levels = var1_levels)
      
      # Sort var2 levels by frequency
      var2_freq <- table(var2_data)
      if(sort_order == "asc") {
        var2_levels <- names(var2_freq)[order(var2_freq)]
      } else {
        var2_levels <- names(var2_freq)[order(var2_freq, decreasing = TRUE)]
      }
      var2_data <- factor(var2_data, levels = var2_levels)
    }
    
    crosstab <- table(var1_data, var2_data, useNA = "no")
    
    # Calculate various percentages
    row_totals <- rowSums(crosstab)
    col_totals <- colSums(crosstab)
    grand_total <- sum(crosstab)
    
    row_prop <- prop.table(crosstab, 1) * 100
    col_prop <- prop.table(crosstab, 2) * 100
    total_prop <- prop.table(crosstab) * 100
    
    # Create the main table structure
    result_list <- list()
    
    # Create header row
    header_row <- data.frame(
      Variable = paste0(var1, " By ", var2),
      stringsAsFactors = FALSE
    )
    
    # Add column headers
    for(col_name in colnames(crosstab)) {
      header_row[[col_name]] <- col_name
    }
    
    # Add Total column if requested
    if("include_totals" %in% stats_options) {
      header_row$Total <- "Total"
    }
    
    result_list[["header"]] <- header_row
    
    # Create statistics header row
    stats_header <- data.frame(Variable = "")
    
    # Build statistics label based on selected options
    stats_label_parts <- c()
    if("count" %in% stats_options) stats_label_parts <- c(stats_label_parts, "Count")
    if("total_percent" %in% stats_options) stats_label_parts <- c(stats_label_parts, "Total %")
    if("col_percent" %in% stats_options) stats_label_parts <- c(stats_label_parts, "Col %")
    if("row_percent" %in% stats_options) stats_label_parts <- c(stats_label_parts, "Row %")
    
    stats_label <- paste(stats_label_parts, collapse = ",")
    
    for(col_name in colnames(crosstab)) {
      stats_header[[col_name]] <- stats_label
    }
    
    if("include_totals" %in% stats_options) {
      stats_header$Total <- stats_label
    }
    
    result_list[["stats_header"]] <- stats_header
    
    # Create data rows
    for(i in 1:nrow(crosstab)) {
      row_name <- rownames(crosstab)[i]
      row_data <- data.frame(Variable = row_name)
      
      for(j in 1:ncol(crosstab)) {
        cell_values <- c()
        
        if("count" %in% stats_options) {
          cell_values <- c(cell_values, crosstab[i, j])
        }
        if("total_percent" %in% stats_options) {
          cell_values <- c(cell_values, round(total_prop[i, j], 2))
        }
        if("col_percent" %in% stats_options) {
          cell_values <- c(cell_values, round(col_prop[i, j], 2))
        }
        if("row_percent" %in% stats_options) {
          cell_values <- c(cell_values, round(row_prop[i, j], 2))
        }
        
        row_data[[colnames(crosstab)[j]]] <- paste(cell_values, collapse = ",")
      }
      
      # Add row total if requested
      if("include_totals" %in% stats_options) {
        total_values <- c()
        if("count" %in% stats_options) {
          total_values <- c(total_values, row_totals[i])
        }
        if("total_percent" %in% stats_options) {
          total_values <- c(total_values, round((row_totals[i] / grand_total) * 100, 2))
        }
        if("col_percent" %in% stats_options) {
          total_values <- c(total_values, "") # No column percentage for row totals
        }
        if("row_percent" %in% stats_options) {
          total_values <- c(total_values, 100) # Row percentage is always 100 for row totals
        }
        
        row_data$Total <- paste(total_values, collapse = ",")
      }
      
      result_list[[paste0("row_", i)]] <- row_data
    }
    
    # Add column totals row if requested
    if("include_totals" %in% stats_options) {
      total_row <- data.frame(Variable = "Total")
      
      for(j in 1:ncol(crosstab)) {
        total_values <- c()
        
        if("count" %in% stats_options) {
          total_values <- c(total_values, col_totals[j])
        }
        if("total_percent" %in% stats_options) {
          total_values <- c(total_values, round((col_totals[j] / grand_total) * 100, 2))
        }
        if("col_percent" %in% stats_options) {
          total_values <- c(total_values, 100) # Column percentage is always 100 for column totals
        }
        if("row_percent" %in% stats_options) {
          total_values <- c(total_values, "") # No row percentage for column totals
        }
        
        total_row[[colnames(crosstab)[j]]] <- paste(total_values, collapse = ",")
      }
      
      # Add grand total
      if("include_totals" %in% stats_options) {
        grand_total_values <- c()
        if("count" %in% stats_options) {
          grand_total_values <- c(grand_total_values, grand_total)
        }
        if("total_percent" %in% stats_options) {
          grand_total_values <- c(grand_total_values, 100)
        }
        if("col_percent" %in% stats_options) {
          grand_total_values <- c(grand_total_values, "")
        }
        if("row_percent" %in% stats_options) {
          grand_total_values <- c(grand_total_values, "")
        }
        
        total_row$Total <- paste(grand_total_values, collapse = ",")
      }
      
      result_list[["total_row"]] <- total_row
    }
    
    # Add chi-square test if requested
    if("chisq" %in% stats_options) {
      chisq_test <- chisq.test(crosstab)
      chi_square_row <- data.frame(
        Variable = paste0("χ²(", chisq_test$parameter, ") = ", 
                          round(chisq_test$statistic, 3), 
                          ", p = ", round(chisq_test$p.value, 4)),
        stringsAsFactors = FALSE
      )
      
      # Add empty columns for the rest
      for(col_name in colnames(crosstab)) {
        chi_square_row[[col_name]] <- ""
      }
      
      if("include_totals" %in% stats_options) {
        chi_square_row$Total <- ""
      }
      
      result_list[["chisq"]] <- chi_square_row
    }
    
    # Combine all rows
    result <- bind_rows(result_list)
    
    return(result)
  }
  
  # Render APA table
  output$apa_table_output <- renderUI({
    table_data <- apa_table_reactive()
    
    if(is.null(table_data)) {
      return(
        tags$div(
          class = "alert alert-info",
          icon("info-circle"),
          "Please select variables in the 'Variable Selection' tab. The table will generate automatically."
        )
      )
    }
    
    # Check if it's an error message
    if("Message" %in% names(table_data)) {
      return(
        tags$div(
          class = "alert alert-warning",
          icon("exclamation-triangle"),
          table_data$Message[1]
        )
      )
    }
    
    # Create flextable
    ft <- flextable(table_data)
    
    # Apply formatting
    ft <- ft %>% 
      font(fontname = "Times New Roman", part = "all") %>%
      fontsize(size = 12, part = "all") %>%
      bold(part = "header") %>%
      flextable::align(align = "center", part = "header") %>%
      flextable::align(j = 1, align = "left", part = "body") %>%
      border_outer() %>%
      border_inner()
    
    # Apply different styles based on selection
    if(input$table_style == "apa") {
      ft <- ft %>% 
        fontsize(size = 12, part = "all")
    } else if(input$table_style == "modern") {
      ft <- ft %>%
        bg(bg = "#f8f9fa", part = "header") %>%
        color(color = "#2C3E50", part = "header") %>%
        bold(part = "header")
    } else if(input$table_style == "classic") {
      ft <- ft %>%
        bg(bg = "#34495E", part = "header") %>%
        color(color = "white", part = "header")
    }
    
    # Convert to HTML for display
    html_output <- htmltools_value(ft)
    
    # Add caption
    caption_text <- ifelse(nchar(input$table_caption) > 0, 
                           input$table_caption,
                           paste("Table", input$table_number))
    
    caption_html <- tags$p(
      style = "font-family: 'Times New Roman'; font-size: 12pt; text-align: center; font-weight: bold; margin-bottom: 20px;",
      caption_text
    )
    
    return(tagList(caption_html, html_output))
  })
  
  # Download Word document
  output$download_doc <- downloadHandler(
    filename = function() {
      paste0("APA_Table_", input$table_number, "_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".docx")
    },
    content = function(file) {
      table_data <- apa_table_reactive()
      req(table_data)
      
      # Don't download if it's an error message
      if("Message" %in% names(table_data)) {
        showNotification("Cannot download - table contains error message", type = "error")
        return()
      }
      
      # Create flextable
      ft <- flextable(table_data)
      
      # Apply formatting
      ft <- ft %>% 
        font(fontname = "Times New Roman", part = "all") %>%
        fontsize(size = 12, part = "all") %>%
        bold(part = "header") %>%
        flextable::align(align = "center", part = "header") %>%
        flextable::align(j = 1, align = "left", part = "body") %>%
        border_outer() %>%
        border_inner() %>%
        width(width = 1.2) %>%
        height(height = 0.3)
      
      # Create Word document
      doc <- officer::read_docx() 
      
      # Add caption
      caption_text <- ifelse(nchar(input$table_caption) > 0, 
                             input$table_caption,
                             paste("Table", input$table_number))
      
      doc <- body_add_par(doc, caption_text, style = "Normal") 
      
      # Add table
      doc <- body_add_flextable(doc, ft)
      
      print(doc, target = file)
      showNotification("Word document downloaded successfully!", type = "message")
    }
  )
}

# Run the application
shinyApp(ui = ui, server = server)
