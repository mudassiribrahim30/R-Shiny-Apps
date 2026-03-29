library(shiny)
library(shinydashboard)
library(DT)
library(haven)
library(readxl)
library(ggplot2)
library(officer)
library(rmarkdown)
library(tidyverse)
library(sortable)
library(scales)

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

# Helper function to format p-values
format_p_value <- function(p_value) {
  if (is.na(p_value)) return("NA")
  if (p_value < 0.001) {
    return("p < 0.001")
  } else {
    return(paste0("p = ", format(round(p_value, 3), nsmall = 3)))
  }
}

# Helper function to perform Pearson's Chi-square test
perform_pearson_chisquare <- function(cont_table) {
  # Perform standard Pearson's Chi-square test
  chisq_result <- chisq.test(cont_table, correct = FALSE)
  return(list(
    statistic = chisq_result$statistic,
    p.value = chisq_result$p.value,
    parameter = chisq_result$parameter
  ))
}

# Define UI
ui <- dashboardPage(
  skin = "blue",
  dashboardHeader(
    title = "CATrend Analyzer",
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
        
        /* Crosstab styling */
        .crosstab-table {
          width: 100%;
          border-collapse: collapse;
          margin: 10px 0;
        }
        
        .crosstab-table th, .crosstab-table td {
          border: 1px solid #ddd;
          padding: 8px;
          text-align: center;
        }
        
        .crosstab-table th {
          background-color: #f2f2f2;
          font-weight: bold;
        }
        
        .crosstab-input {
          width: 80px;
          text-align: center;
        }
        
        .data-type-card {
          cursor: pointer;
          transition: all 0.3s ease;
          border: 2px solid transparent;
        }
        
        .data-type-card:hover {
          transform: translateY(-5px);
          box-shadow: 0 6px 12px rgba(0,0,0,0.15);
        }
        
        .data-type-card.selected {
          border-color: #0093D0;
          background-color: #f0f8ff;
        }
        
        /* Plot customization controls */
        .plot-controls {
          background: #f8f9fa;
          padding: 15px;
          border-radius: 5px;
          margin-bottom: 15px;
          border-left: 4px solid #17a2b8;
        }
        
        /* Professional font styling for test results */
        .test-results {
          font-family: 'Courier New', monospace;
          font-size: 14px;
          line-height: 1.4;
          white-space: pre-wrap;
        }
        
        .test-results-header {
          font-weight: bold;
          font-size: 16px;
          margin-bottom: 10px;
        }
        
        /* Enhanced table styling */
        .enhanced-table {
          font-size: 14px;
        }
        
        .enhanced-table th {
          background-color: #f8f9fa;
          font-weight: bold;
          text-align: center;
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
                           p("The application uses this formula to compute the Cochran–Armitage test for trend:"),
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
                           p("Written by:"),
                           tags$h4("Mudasir Mohammed Ibrahim"),
                           
                           # Academic profile links with icons
                           div(style = "margin: 15px 0; display: flex; justify-content: center; gap: 20px; flex-wrap: wrap;",
                               
                               # ResearchGate
                               a(href = "https://www.researchgate.net/profile/Mudasir-Ibrahim", 
                                 target = "_blank",
                                 title = "ResearchGate Profile",
                                 style = "color: #2d3748; text-decoration: none; display: inline-flex; flex-direction: column; align-items: center; transition: transform 0.3s ease;",
                                 onmouseover = "this.style.transform='translateY(-3px)'",
                                 onmouseout = "this.style.transform='translateY(0)'",
                                 div(style = "background: #00ccbb; width: 45px; height: 45px; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin-bottom: 5px; box-shadow: 0 4px 10px rgba(0,0,0,0.1);",
                                     icon("researchgate", style = "font-size: 22px; color: white;")
                                 ),
                                 div(style = "font-size: 11px; font-weight: 500;", "ResearchGate")
                               ),
                               
                               
                               # Web of Science / Publons
                               a(href = "https://www.webofscience.com/wos/author/record/HPC-2085-2023", 
                                 target = "_blank",
                                 title = "Web of Science / Publons Profile",
                                 style = "color: #2d3748; text-decoration: none; display: inline-flex; flex-direction: column; align-items: center; transition: transform 0.3s ease;",
                                 onmouseover = "this.style.transform='translateY(-3px)'",
                                 onmouseout = "this.style.transform='translateY(0)'",
                                 div(style = "background: #ff6b6b; width: 45px; height: 45px; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin-bottom: 5px; box-shadow: 0 4px 10px rgba(0,0,0,0.1);",
                                     icon("book", style = "font-size: 22px; color: white;")
                                 ),
                                 div(style = "font-size: 11px; font-weight: 500;", "Web of Science")
                               ),
                               
                               # GitHub
                               a(href = "https://github.com/mudassiribrahim30", 
                                 target = "_blank",
                                 title = "GitHub Profile",
                                 style = "color: #2d3748; text-decoration: none; display: inline-flex; flex-direction: column; align-items: center; transition: transform 0.3s ease;",
                                 onmouseover = "this.style.transform='translateY(-3px)'",
                                 onmouseout = "this.style.transform='translateY(0)'",
                                 div(style = "background: #333; width: 45px; height: 45px; border-radius: 50%; display: flex; align-items: center; justify-content: center; margin-bottom: 5px; box-shadow: 0 4px 10px rgba(0,0,0,0.1);",
                                     icon("github", style = "font-size: 22px; color: white;")
                                 ),
                                 div(style = "font-size: 11px; font-weight: 500;", "GitHub")
                               )
                           ),
                           
                           # ORCID (optional - can add if you have one)
                           tags$p("ORCID: ", tags$a(href="https://orcid.org/0000-0002-9049-8222", 
                                                    target = "_blank",
                                                    style = "color: #38b2ac; font-weight: 500;",
                                                    "0000-0002-9049-8222"))
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
      
      # Data Upload Tab - UPDATED WITH TWO OPTIONS
      tabItem(tabName = "upload",
              fluidRow(
                box(
                  title = "Select Data Type", status = "primary", solidHeader = TRUE, width = 12,
                  fluidRow(
                    column(6,
                           div(class = "data-type-card",
                               style = "padding: 20px; text-align: center; background: white; border-radius: 8px; cursor: pointer;",
                               onclick = "$('#data_type_normal').click();",
                               h4(icon("table"), "Normal Data"),
                               p("Raw dataset with individual observations"),
                               p("Suitable for standard analysis"),
                               radioButtons("data_type", NULL, 
                                            choices = c("Normal Data" = "normal"), 
                                            selected = "normal")
                           )
                    ),
                    column(6,
                           div(class = "data-type-card",
                               style = "padding: 20px; text-align: center; background: white; border-radius: 8px; cursor: pointer;",
                               onclick = "$('#data_type_crosstab').click();",
                               h4(icon("th"), "Crosstabulation Data"),
                               p("Already summarized in contingency table format"),
                               p("Enter counts manually"),
                               radioButtons("data_type", NULL, 
                                            choices = c("Crosstabulation Data" = "crosstab"), 
                                            selected = character(0))
                           )
                    )
                  )
                )
              ),
              
              # Normal Data Upload Section
              conditionalPanel(
                condition = "input.data_type == 'normal'",
                fluidRow(
                  box(
                    title = "Upload Data File", status = "info", solidHeader = TRUE, width = 12,
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
                )
              ),
              
              # Crosstabulation Data Section
              conditionalPanel(
                condition = "input.data_type == 'crosstab'",
                fluidRow(
                  box(
                    title = "Crosstabulation Data Setup", status = "info", solidHeader = TRUE, width = 12,
                    fluidRow(
                      column(6,
                             textInput("dep_var_name", "Dependent Variable Name:", value = "Outcome"),
                             textInput("ind_var_name", "Independent Variable Name:", value = "Exposure"),
                             numericInput("n_dep_levels", "Number of Dependent Variable Levels:", 
                                          value = 2, min = 2, max = 10),
                             numericInput("n_ind_levels", "Number of Independent Variable Levels:", 
                                          value = 3, min = 2, max = 10),
                             actionButton("setup_crosstab", "Setup Crosstabulation Table", 
                                          class = "btn-primary")
                      ),
                      column(6,
                             div(class = "instruction-highlight",
                                 icon("info-circle"),
                                 "Set up your contingency table by specifying the number of levels for both variables, then enter the counts in the table that will appear below."
                             )
                      )
                    )
                  )
                ),
                
                # Dynamic crosstabulation table UI
                uiOutput("crosstab_ui"),
                
                # Action button to finalize crosstab data
                fluidRow(
                  box(
                    title = "Finalize Crosstabulation Data", status = "success", solidHeader = TRUE, width = 12,
                    actionButton("create_crosstab_data", "Create Dataset from Crosstabulation", 
                                 class = "btn-success", icon = icon("check")),
                    helpText("Click this button after entering all counts to create the dataset for analysis.")
                  )
                )
              ),
              
              # Data Preview (common to both types)
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
                      checkboxInput("continuity_correction", "Apply Continuity Correction", value = FALSE),
                      actionButton("analyze", "Run Cochran-Armitage Test", 
                                   class = "btn-primary", icon = icon("play")),
                      helpText("Note: Analysis will use the cleaned data from the Data Cleaning tab. The test type (increasing or decreasing) is automatically determined based on the sign of the Z-statistic.")
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
              
              # Enhanced Crosstabulation Results with Percentages
              fluidRow(
                box(
                  title = "Enhanced Crosstabulation Results", status = "info", solidHeader = TRUE, width = 12,
                  div(class = "results-panel",
                      h4("Contingency Table with Row, Column, and Total Percentages"),
                      DTOutput("enhanced_crosstab")
                  )
                )
              ),
              
              # Enhanced Plot Controls
              fluidRow(
                box(
                  title = "Plot Customization", status = "warning", solidHeader = TRUE, width = 12,
                  div(class = "plot-controls",
                      fluidRow(
                        # In the Plot Customization section of the Analysis tab
                        column(3,
                               selectInput("plot_type", "Plot Value Type:",
                                           choices = c("Counts" = "counts",
                                                       "Row Percentages" = "row_percent",
                                                       "Column Percentages" = "col_percent"),
                                           selected = "col_percent")
                        ),
                        column(3,
                               textInput("plot_title", "Plot Title:", 
                                         value = "Stacked Bar Plot"),
                               numericInput("title_size", "Title Size:", 
                                            value = 16, min = 8, max = 24)
                        ),
                        column(3,
                               numericInput("label_size", "Label Size:", 
                                            value = 12, min = 6, max = 20),
                               numericInput("legend_size", "Legend Text Size:", 
                                            value = 12, min = 6, max = 20)
                        ),
                        column(3,
                               numericInput("axis_title_size", "Axis Title Size:", 
                                            value = 14, min = 8, max = 20),
                               numericInput("axis_text_size", "Axis Text Size:", 
                                            value = 12, min = 6, max = 18)
                        )
                      ),
                      fluidRow(
                        column(12,
                               h5("Custom Colors for Dependent Variable Levels:"),
                               uiOutput("color_pickers_ui")
                        )
                      )
                  )
                )
              ),
              
              # Enhanced Stacked Bar Plot
              fluidRow(
                box(
                  title = "Enhanced Stacked Bar Plot", status = "info", solidHeader = TRUE, width = 12,
                  plotOutput("enhanced_bar_plot", height = "600px"),
                  
                  # ADD THIS DOWNLOAD INSTRUCTION SECTION
                  div(
                    style = "margin-top: 20px; padding: 15px; background-color: #FFF9C4; border-left: 4px solid #F39C12; border-radius: 4px;",
                    tags$div(
                      style = "display: flex; align-items: center;",
                      tags$strong(style = "color: #7D6608; font-size: 16px;", 
                                  icon("download"), "DOWNLOAD FULL ANALYSIS REPORT:"),
                      tags$span(style = "margin-left: 10px; color: #7D6608; font-weight: bold;",
                                "Go to the 'Results' section and click on 'Download Full Report' to get the complete analysis including this graph plot, statistical results, and detailed interpretation.")
                    )
                  )
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
  crosstab_data <- reactiveVal(NULL)
  custom_colors <- reactiveVal(list())
  
  # Track if analysis has been completed
  output$analysis_done <- reactive({
    analysis_done()
  })
  outputOptions(output, "analysis_done", suspendWhenHidden = FALSE)
  
  # Load data based on file type - UPDATED FOR SPSS AND STATA
  observe({
    tryCatch({
      req(input$file_type, input$data_type == "normal")
      
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
  
  # Crosstabulation UI
  output$crosstab_ui <- renderUI({
    req(input$setup_crosstab, input$n_dep_levels, input$n_ind_levels)
    
    n_dep <- input$n_dep_levels
    n_ind <- input$n_ind_levels
    
    fluidRow(
      box(
        title = "Crosstabulation Table - Enter Counts", status = "primary", solidHeader = TRUE, width = 12,
        div(
          style = "overflow-x: auto;",
          tags$table(
            class = "crosstab-table",
            tags$tr(
              tags$th(""),
              tags$th(colspan = n_ind, input$ind_var_name)
            ),
            tags$tr(
              tags$th(paste(input$dep_var_name, "↓ /", input$ind_var_name, "→")),
              lapply(1:n_ind, function(j) {
                tags$th(textInput(paste0("ind_level_", j), paste("Level", j), value = paste("Level", j)))
              })
            ),
            lapply(1:n_dep, function(i) {
              tags$tr(
                tags$th(textInput(paste0("dep_level_", i), paste("Level", i), value = paste("Level", i))),
                lapply(1:n_ind, function(j) {
                  tags$td(
                    numericInput(paste0("count_", i, "_", j), NULL, value = 0, min = 0, step = 1,
                                 width = "80px")
                  )
                })
              )
            })
          )
        )
      )
    )
  })
  
  # Create dataset from crosstabulation
  observeEvent(input$create_crosstab_data, {
    req(input$n_dep_levels, input$n_ind_levels)
    
    tryCatch({
      n_dep <- input$n_dep_levels
      n_ind <- input$n_ind_levels
      
      # Get level names
      dep_levels <- sapply(1:n_dep, function(i) {
        input[[paste0("dep_level_", i)]]
      })
      
      ind_levels <- sapply(1:n_ind, function(j) {
        input[[paste0("ind_level_", j)]]
      })
      
      # Get counts and create dataset
      observations <- list()
      for (i in 1:n_dep) {
        for (j in 1:n_ind) {
          count <- input[[paste0("count_", i, "_", j)]]
          if (!is.null(count) && count > 0) {
            # Create observations for each count
            for (k in 1:count) {
              observations[[length(observations) + 1]] <- data.frame(
                dep_var = dep_levels[i],
                ind_var = ind_levels[j],
                stringsAsFactors = FALSE
              )
            }
          }
        }
      }
      
      if (length(observations) > 0) {
        df <- do.call(rbind, observations)
        names(df) <- c(input$dep_var_name, input$ind_var_name)
        
        data_loaded(df)
        data_cleaned(df)
        crosstab_data(TRUE)
        
        showNotification(paste("Crosstabulation dataset created with", nrow(df), "observations"), 
                         type = "message")
      } else {
        showNotification("Please enter some counts greater than 0", type = "warning")
      }
    }, error = function(e) {
      showNotification(paste("Error creating dataset:", e$message), type = "error")
    })
  })
  
  # Data preview
  output$data_preview <- renderDT({
    req(data_loaded())
    datatable(data_loaded(), 
              options = list(scrollX = TRUE, pageLength = 5),
              caption = if (isTRUE(crosstab_data())) {
                "Crosstabulation Data (Created from Counts)"
              } else {
                "Original Data Preview"
              })
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
  
  # Color pickers for dependent variable levels
  output$color_pickers_ui <- renderUI({
    req(data_cleaned(), input$dependent_var)
    
    dep_levels <- unique(data_cleaned()[[input$dependent_var]])
    dep_levels <- dep_levels[!is.na(dep_levels)]
    current_dep_order <- if(!is.null(dep_level_order())) dep_level_order() else as.character(dep_levels)
    
    # Default colors
    default_colors <- c("#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd", 
                        "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf")
    
    tagList(
      fluidRow(
        lapply(seq_along(current_dep_order), function(i) {
          level <- current_dep_order[i]
          default_color <- default_colors[((i-1) %% length(default_colors)) + 1]
          column(3,
                 colourpicker::colourInput(
                   inputId = paste0("color_", gsub("[^a-zA-Z0-9]", "_", level)),
                   label = level,
                   value = default_color
                 )
          )
        })
      )
    )
  })
  
  # Get custom colors
  observe({
    req(data_cleaned(), input$dependent_var)
    
    dep_levels <- unique(data_cleaned()[[input$dependent_var]])
    dep_levels <- dep_levels[!is.na(dep_levels)]
    current_dep_order <- if(!is.null(dep_level_order())) dep_level_order() else as.character(dep_levels)
    
    colors_list <- list()
    for (level in current_dep_order) {
      color_id <- paste0("color_", gsub("[^a-zA-Z0-9]", "_", level))
      if (!is.null(input[[color_id]])) {
        colors_list[[level]] <- input[[color_id]]
      }
    }
    
    if (length(colors_list) > 0) {
      custom_colors(colors_list)
    }
  })
  
  # ACCURATE Cochran-Armitage test implementation using the provided formula
  perform_cochran_armitage_accurate <- function(cont_table, test_type = "two.sided", continuity_correction = FALSE) {
    # Implementation following the theoretical background provided
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
    
    # Create comprehensive result object
    result <- list(
      statistic = c("Z" = z),
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
  
  # Enhanced crosstabulation with percentages - CORRECTED VERSION
  output$enhanced_crosstab <- renderDT({
    req(analysis_results())
    results <- analysis_results()
    
    cont_table <- results$contingency_table
    
    # Calculate percentages - CORRECTED: swapped row and column calculations
    row_pct <- prop.table(cont_table, 1) * 100  # Percentage within each ROW (dependent variable level)
    col_pct <- prop.table(cont_table, 2) * 100  # Percentage within each COLUMN (independent variable level)
    total_pct <- prop.table(cont_table) * 100   # Percentage of total
    
    # Create combined table
    row_names <- rownames(cont_table)
    col_names <- colnames(cont_table)
    
    # Combine all tables - CORRECTED LABELS
    combined_table <- data.frame(
      Measure = c("Counts", rep("", nrow(cont_table)-1),
                  "Column % (Within Dependent Level)", rep("", nrow(cont_table)-1),
                  "Row % (Within Independent Level)", rep("", nrow(cont_table)-1),
                  "Total %", rep("", nrow(cont_table)-1)),
      Row = rep(row_names, 4)
    )
    
    # Add data for each measure type - CORRECTED: using proper percentages
    for (j in 1:ncol(cont_table)) {
      col_name <- col_names[j]
      combined_table[[col_name]] <- c(
        # Counts
        cont_table[, j],
        # Row percentages (percentage within each dependent variable level)
        sprintf("%.1f%%", row_pct[, j]),
        # Column percentages (percentage within each independent variable level)  
        sprintf("%.1f%%", col_pct[, j]),
        # Total percentages
        sprintf("%.1f%%", total_pct[, j])
      )
    }
    
    datatable(
      combined_table,
      options = list(
        scrollX = TRUE,
        pageLength = 20,
        dom = 'Bfrtip',
        buttons = c('copy', 'csv', 'excel', 'pdf')
      ),
      caption = paste("Enhanced Crosstabulation: Counts, Row %, Column %, and Total %"),
      rownames = FALSE
    ) %>%
      formatStyle(
        columns = 1:ncol(combined_table),
        fontSize = '14px'
      )
  })
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
      
      # --- PERFORM PEARSON'S CHI-SQUARE TEST ---
      pearson_chisq <- perform_pearson_chisquare(cont_table)
      
      # --- FIRST RUN TWO-SIDED TEST TO GET Z-STATISTIC SIGN ---
      # Run a two-sided test first to determine the direction
      two_sided_test <- perform_cochran_armitage_accurate(
        cont_table, 
        test_type = "two.sided",
        continuity_correction = input$continuity_correction
      )
      
      z_stat <- two_sided_test$statistic["Z"]
      
      # Determine the appropriate one-sided test type based on Z-statistic sign
      if (z_stat > 0) {
        test_type <- "one.sided.increasing"
      } else if (z_stat < 0) {
        test_type <- "one.sided.decreasing"
      } else {
        test_type <- "two.sided"
      }
      
      # Perform the actual test with the determined test type
      ca_test <- perform_cochran_armitage_accurate(
        cont_table, 
        test_type = test_type,
        continuity_correction = input$continuity_correction
      )
      
      # Calculate proportions for display
      first_dep_level <- rownames(cont_table)[1]
      first_dep_counts <- cont_table[first_dep_level, ]
      row_totals_first <- rowSums(cont_table)[first_dep_level]
      proportions <- first_dep_counts / row_totals_first
      
      # Determine trend direction for interpretation based on Z-statistic
      is_increasing <- z_stat > 0
      is_decreasing <- z_stat < 0
      
      # Create enhanced plot data
      plot_data <- df_processed %>%
        count(!!sym(ind_var), !!sym(dep_var)) %>%
        group_by(!!sym(ind_var)) %>%
        mutate(
          count = n,
          row_percent = n / sum(n) * 100,
          col_percent = n / sum(n[!!sym(ind_var) == first(!!sym(ind_var))]) * 100,
          total_percent = n / sum(n) * 100
        ) %>%
        ungroup()
      
      # Set analysis as done
      analysis_done(TRUE)
      
      list(
        contingency_table = cont_table,
        test_results = ca_test,
        two_sided_test = two_sided_test,
        pearson_chisq = pearson_chisq,
        plot_data = plot_data,
        data = df_processed,
        dep_var = dep_var,
        ind_var = ind_var,
        dep_levels = levels(df_processed[[dep_var]]),
        ind_levels = levels(df_processed[[ind_var]]),
        sample_size = nrow(df_processed),
        original_sample_size = nrow(df),
        proportions = proportions,
        is_increasing = is_increasing,
        is_decreasing = is_decreasing,
        test_type_used = test_type,
        first_dep_level = first_dep_level,
        z_stat = z_stat
      )
    }, error = function(e) {
      showNotification(paste("Analysis error:", e$message), type = "error")
      return(NULL)
    })
  })
  
  # Enhanced bar plot - CORRECTED VERSION WITH PROPER PERCENTAGE CALCULATIONS
  output$enhanced_bar_plot <- renderPlot({
    req(analysis_results(), input$plot_type)
    
    results <- analysis_results()
    
    # Get the processed data and create proper plot data
    df_processed <- results$data
    dep_var <- results$dep_var
    ind_var <- results$ind_var
    
    # Create fresh plot data with proper calculations
    plot_data <- df_processed %>%
      count(!!sym(ind_var), !!sym(dep_var)) %>%
      group_by(!!sym(ind_var)) %>%
      mutate(
        row_percent = n / sum(n) * 100  # Row percentage: within each independent variable group
      ) %>%
      ungroup() %>%
      group_by(!!sym(dep_var)) %>%
      mutate(
        col_percent = n / sum(n) * 100  # Column percentage: within each dependent variable level
      ) %>%
      ungroup() %>%
      mutate(
        total_percent = n / sum(n) * 100  # Total percentage: of overall total
      )
    
    # Determine which value to plot based on user selection - CORRECTED LOGIC
    if (input$plot_type == "counts") {
      plot_data$plot_value <- plot_data$n
      y_label <- "Count"
      label_format <- function(x) format(x, big.mark = ",")
    } else if (input$plot_type == "row_percent") {
      plot_data$plot_value <- plot_data$row_percent
      y_label <- "Percentage Within Group (%)"
      label_format <- function(x) paste0(format(round(x, 1), nsmall = 1), "%")
    } else if (input$plot_type == "col_percent") {
      plot_data$plot_value <- plot_data$col_percent
      y_label <- "Percentage Within Category (%)"
      label_format <- function(x) paste0(format(round(x, 1), nsmall = 1), "%")
    }
    
    # Get custom colors
    colors <- custom_colors()
    if (length(colors) == 0) {
      # Use default colors if no custom colors set
      colors <- scales::hue_pal()(length(results$dep_levels))
      names(colors) <- results$dep_levels
    }
    
    # Create the enhanced plot with professional styling
    p <- ggplot(plot_data, aes(x = !!sym(ind_var), y = plot_value, 
                               fill = !!sym(dep_var))) +
      geom_bar(stat = "identity", position = "stack") +
      geom_text(aes(label = label_format(plot_value)), 
                position = position_stack(vjust = 0.5), 
                size = input$label_size, 
                color = "white",
                fontface = "bold") +
      labs(
        title = input$plot_title,
        x = ind_var,
        y = y_label,
        fill = dep_var
      ) +
      theme_minimal() +
      theme(
        plot.title = element_text(size = input$title_size, face = "bold", hjust = 0.5),
        axis.title = element_text(size = input$axis_title_size, face = "bold"),
        axis.text = element_text(size = input$axis_text_size),
        axis.text.x = element_text(angle = 45, hjust = 1),
        legend.title = element_text(size = input$legend_size, face = "bold"),
        legend.text = element_text(size = input$legend_size),
        legend.position = "bottom",
        panel.grid.major = element_line(color = "grey80", linewidth = 0.2),
        panel.grid.minor = element_blank(),
        plot.background = element_rect(fill = "white", color = NA),
        panel.background = element_rect(fill = "white", color = NA)
      )
    
    # Apply custom colors if available
    if (length(colors) > 0) {
      p <- p + scale_fill_manual(values = colors)
    }
    
    # Adjust y-axis for percentages
    if (input$plot_type %in% c("row_percent", "col_percent")) {
      p <- p + scale_y_continuous(labels = function(x) paste0(format(round(x, 1), nsmall = 1), "%"))
    } else {
      p <- p + scale_y_continuous(labels = label_format)
    }
    
    return(p)
  })
  # Test results - CONSISTENT OUTPUT WITH FORMATTED P-VALUES AND CORRECT CHI-SQUARE
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
    cat("- Continuity correction:", ifelse(results$test_results$continuity_correction, "Yes", "No"), "\n\n")
    
    cat("LEVEL ORDERING:\n")
    cat("- Dependent:", paste(results$dep_levels, collapse = " -> "), "\n")
    cat("- Independent:", paste(results$ind_levels, collapse = " < "), "\n\n")
    
    cat("CONTINGENCY TABLE:\n")
    print(results$contingency_table)
    cat("\n")
    
    cat("PROPORTION ANALYSIS:\n")
    # Get the first level of dependent var for proportion calculation
    first_dep_level <- results$first_dep_level
    first_dep_counts <- results$contingency_table[first_dep_level, ]
    row_totals_first <- rowSums(results$contingency_table)[first_dep_level]
    proportions <- first_dep_counts / row_totals_first
    
    cat("- First dependent level selected for proportion calculation:", first_dep_level, "\n")
    cat("- Proportions across independent levels:", paste(format(round(proportions, 4), nsmall = 4), collapse = " -> "), "\n")
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
    
    # Extract statistics
    z_stat <- results$test_results$statistic["Z"]
    one_sided_p <- results$test_results$p.value
    two_sided_p <- results$two_sided_test$p.value
    
    # Pearson's Chi-square test results
    pearson_chi_sq <- results$pearson_chisq$statistic
    pearson_df <- results$pearson_chisq$parameter
    pearson_p <- results$pearson_chisq$p.value
    
    # Display Cochran-Armitage results
    cat("\nCochran-Armitage Trend Test:\n")
    cat("Z-statistic:", format(z_stat, digits = 4), "\n\n")
    
    # Display Pearson's Chi-square test results
    cat("Pearson's Chi-square Test of Independence:\n")
    cat("Chi-square statistic:", format(pearson_chi_sq, digits = 4), "\n")
    cat("Degrees of freedom:", pearson_df, "\n")
    cat("p-value:", format_p_value(pearson_p), "\n\n")
    
    # Display appropriate p-values based on Z-statistic sign
    if (results$is_increasing) {
      cat("Cochran-Armitage One-sided Test (Increasing Trend):", format_p_value(one_sided_p), "\n")
      cat("Cochran-Armitage Two-sided p-value:", format_p_value(two_sided_p), "\n")
    } else if (results$is_decreasing) {
      cat("Cochran-Armitage One-sided Test (Decreasing Trend):", format_p_value(one_sided_p), "\n")
      cat("Cochran-Armitage Two-sided p-value:", format_p_value(two_sided_p), "\n")
    } else {
      cat("Cochran-Armitage Two-sided p-value:", format_p_value(two_sided_p), "\n")
    }
    cat("\n")
    
    cat("INTERPRETATION:\n")
    # Use one-sided p for trend interpretation when direction is clear
    p_val_interpret <- if (results$is_increasing || results$is_decreasing) one_sided_p else two_sided_p
    
    if (p_val_interpret < 0.05) {
      cat("✓ Statistically significant trend detected (", format_p_value(p_val_interpret), ")\n")
      if (results$is_increasing) {
        cat("  Evidence supports an INCREASING trend in the proportion of '", 
            first_dep_level, "' across the ordered levels of ", 
            results$ind_var, " (Z = ", format(z_stat, digits = 3), ").\n", sep = "")
        cat("  As ", results$ind_var, " increases, the probability of '", 
            first_dep_level, "' significantly increases.\n", sep = "")
      } else if (results$is_decreasing) {
        cat("  Evidence supports a DECREASING trend in the proportion of '", 
            first_dep_level, "' across the ordered levels of ", 
            results$ind_var, " (Z = ", format(z_stat, digits = 3), ").\n", sep = "")
        cat("  As ", results$ind_var, " increases, the probability of '", 
            first_dep_level, "' significantly decreases.\n", sep = "")
      } else {
        cat("  Evidence suggests a systematic relationship between", results$dep_var, "and", results$ind_var, "\n")
      }
    } else {
      cat("✗ No statistically significant trend detected (", format_p_value(p_val_interpret), ")\n")
      cat("  No evidence of systematic relationship between", results$dep_var, 
          "and", results$ind_var, "\n")
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
  
  # Report preview
  output$report_preview <- renderUI({
    req(analysis_results())
    
    results <- analysis_results()
    z_stat <- results$test_results$statistic["Z"]
    one_sided_p <- results$test_results$p.value
    two_sided_p <- results$two_sided_test$p.value
    pearson_chi_sq <- results$pearson_chisq$statistic
    pearson_df <- results$pearson_chisq$parameter
    pearson_p <- results$pearson_chisq$p.value
    first_dep_level <- results$first_dep_level
    
    HTML(paste(
      "<h4>Report Preview</h4>",
      "<div style='background-color: #f8f9fa; padding: 15px; border-radius: 5px;'>",
      "<p><strong>Dependent Variable:</strong>", results$dep_var, "</p>",
      "<p><strong>Independent Variable:</strong>", results$ind_var, "</p>",
      "<p><strong>Sample Size:</strong>", results$sample_size, "</p>",
      "<p><strong>Trend Detected:</strong>", 
      if (results$is_increasing) "Increasing" else if (results$is_decreasing) "Decreasing" else "No clear monotonic trend", "</p>",
      "<p><strong>Continuity Correction:</strong>", ifelse(results$test_results$continuity_correction, "Yes", "No"), "</p>",
      "<p><strong>Cochran-Armitage Z-statistic:</strong>", format(z_stat, digits = 4), "</p>",
      "<p><strong>Pearson's Chi-square:</strong> χ² = ", format(pearson_chi_sq, digits = 4), ", df = ", pearson_df, ", ", format_p_value(pearson_p), "</p>",
      if (results$is_increasing) {
        paste0("<p><strong>One-sided p-value (increasing):</strong> ", format_p_value(one_sided_p), "</p>")
      } else if (results$is_decreasing) {
        paste0("<p><strong>One-sided p-value (decreasing):</strong> ", format_p_value(one_sided_p), "</p>")
      } else {
        ""
      },
      "<p><strong>Two-sided p-value:</strong> ", format_p_value(two_sided_p), "</p>",
      "<p><strong>Significance:</strong> <span style='color:", 
      if (one_sided_p < 0.05) "#28a745" else "#dc3545", "'>",
      if (one_sided_p < 0.05) "Statistically Significant" else "Not Significant",
      "</span></p>",
      if (results$sample_size < 30) "<p style='color: #856404;'>⚠️ Note: Small sample size</p>" else "",
      "</div>"
    ))
  })
  
  # Download report - ENHANCED VERSION WITH FIXED PLOT
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
        z_stat <- results$test_results$statistic["Z"]
        one_sided_p <- results$test_results$p.value
        two_sided_p <- results$two_sided_test$p.value
        pearson_chi_sq <- results$pearson_chisq$statistic
        pearson_df <- results$pearson_chisq$parameter
        pearson_p <- results$pearson_chisq$p.value
        cont_table <- results$contingency_table
        dep_var <- results$dep_var
        ind_var <- results$ind_var
        sample_size <- results$sample_size
        continuity_correction <- results$test_results$continuity_correction
        first_dep_level <- results$first_dep_level
        
        # Format p-values for report
        format_p_val <- function(p) {
          if (p < 0.001) return("p < 0.001")
          return(paste0("p = ", format(round(p, 3), nsmall = 3)))
        }
        
        # Calculate enhanced crosstabulation for report
        row_pct <- prop.table(cont_table, 1) * 100
        col_pct <- prop.table(cont_table, 2) * 100
        total_pct <- prop.table(cont_table) * 100
        
        # Save the enhanced plot to a temporary file - USING CURRENT PLOT SETTINGS FROM APP
        temp_plot <- file.path(tempdir(), "enhanced_plot.png")
        
        # Recreate the enhanced plot for the report using CURRENT APP SETTINGS
        plot_data <- results$plot_data
        
        # Use the same plot type as currently selected in the app
        current_plot_type <- input$plot_type
        if (is.null(current_plot_type)) {
          current_plot_type <- "col_percent"  # Default if not set
        }
        
        # Determine plot values based on current app selection - CORRECTED
        if (current_plot_type == "counts") {
          plot_data$plot_value <- plot_data$count
          y_label <- "Count"
          label_format <- function(x) format(x, big.mark = ",")
        } else if (current_plot_type == "row_percent") {
          plot_data$plot_value <- plot_data$row_percent
          y_label <- "Percentage Within Group (%)"
          label_format <- function(x) paste0(format(round(x, 1), nsmall = 1), "%")
        } else if (current_plot_type == "col_percent") {
          plot_data$plot_value <- plot_data$col_percent
          y_label <- "Percentage Within Category (%)"
          label_format <- function(x) paste0(format(round(x, 1), nsmall = 1), "%")
        }
        
        # Get current custom colors from app
        colors <- custom_colors()
        if (length(colors) == 0) {
          colors <- scales::hue_pal()(length(results$dep_levels))
          names(colors) <- results$dep_levels
        }
        
        # Get current plot title from app or use default
        current_plot_title <- ifelse(!is.null(input$plot_title) && input$plot_title != "", 
                                     input$plot_title, 
                                     paste("Stacked Bar Plot:", dep_var, "by", ind_var))
        
        # Create plot with EXACTLY the same styling as the app
        report_plot <- ggplot(plot_data, aes(x = !!sym(ind_var), y = plot_value, 
                                             fill = !!sym(dep_var))) +
          geom_bar(stat = "identity", position = "stack") +
          geom_text(aes(label = label_format(plot_value)), 
                    position = position_stack(vjust = 0.5), 
                    size = 4, 
                    color = "white",
                    fontface = "bold") +
          labs(
            title = current_plot_title,
            x = ind_var,
            y = y_label,
            fill = dep_var
          ) +
          theme_minimal() +
          theme(
            plot.title = element_text(size = 16, face = "bold", hjust = 0.5),
            axis.title = element_text(size = 14, face = "bold"),
            axis.text = element_text(size = 12),
            axis.text.x = element_text(angle = 45, hjust = 1),
            legend.title = element_text(size = 12, face = "bold"),
            legend.text = element_text(size = 12),
            legend.position = "bottom",
            panel.grid.major = element_line(color = "grey80", linewidth = 0.2),
            panel.grid.minor = element_blank(),
            plot.background = element_rect(fill = "white", color = NA),
            panel.background = element_rect(fill = "white", color = NA)
          ) +
          scale_fill_manual(values = colors)
        
        # Adjust y-axis for percentages
        if (current_plot_type %in% c("row_percent", "col_percent")) {
          report_plot <- report_plot + scale_y_continuous(labels = function(x) paste0(format(round(x, 1), nsmall = 1), "%"))
        }
        
        ggsave(temp_plot, plot = report_plot, width = 10, height = 8, dpi = 300)
        
        # Create the Rmd content with enhanced crosstabulation
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
          "- **Trend Detected:** ", 
          if (results$is_increasing) "Increasing" else if (results$is_decreasing) "Decreasing" else "No clear monotonic trend", "\n",
          "- **Continuity Correction:** ", ifelse(continuity_correction, "Yes", "No"), "\n",
          "- **Plot Type:** ", switch(current_plot_type,
                                      "counts" = "Counts",
                                      "row_percent" = "Percentage Within Group (Row %)", 
                                      "col_percent" = "Percentage Within Category (Column %)"), "\n\n",
          
          "## Level Ordering\n\n",
          "- **Dependent Variable Levels:** ", paste(results$dep_levels, collapse = " → "), "\n",
          "- **Independent Variable Levels:** ", paste(results$ind_levels, collapse = " < "), "\n\n",
          
          "## Enhanced Crosstabulation Results\n\n",
          "### Contingency Table with Counts\n",
          "```{r, echo=FALSE}\n",
          "knitr::kable(cont_table, caption = 'Contingency Table - Counts')\n",
          "```\n\n",
          
          "### Row Percentages\n",
          "```{r, echo=FALSE}\n",
          "row_pct_table <- prop.table(cont_table, 1) * 100\n",
          "knitr::kable(round(row_pct_table, 1), caption = 'Row Percentages (%)')\n",
          "```\n\n",
          
          "### Column Percentages\n", 
          "```{r, echo=FALSE}\n",
          "col_pct_table <- prop.table(cont_table, 2) * 100\n",
          "knitr::kable(round(col_pct_table, 1), caption = 'Column Percentages (%)')\n",
          "```\n\n",
          
          "### Total Percentages\n",
          "```{r, echo=FALSE}\n",
          "total_pct_table <- prop.table(cont_table) * 100\n",
          "knitr::kable(round(total_pct_table, 1), caption = 'Total Percentages (%)')\n",
          "```\n\n",
          
          "## Statistical Results\n\n",
          "### Cochran-Armitage Trend Test\n",
          "- **Z-statistic:** ", format(z_stat, digits = 4), "\n",
          if (results$is_increasing) {
            paste0("- **One-sided p-value (increasing):** ", format_p_val(one_sided_p), "\n")
          } else if (results$is_decreasing) {
            paste0("- **One-sided p-value (decreasing):** ", format_p_val(one_sided_p), "\n")
          } else {
            ""
          },
          "- **Two-sided p-value:** ", format_p_val(two_sided_p), "\n\n",
          
          "### Pearson's Chi-square Test of Independence\n",
          "- **Chi-square statistic:** ", format(pearson_chi_sq, digits = 4), "\n",
          "- **Degrees of freedom:** ", pearson_df, "\n",
          "- **p-value:** ", format_p_val(pearson_p), "\n\n",
          
          "## Visualization\n\n",
          "**Plot Type: ", switch(current_plot_type,
                                  "counts" = "Counts",
                                  "row_percent" = "Percentage Within Group (Row %)",
                                  "col_percent" = "Percentage Within Category (Column %)"), "**\n\n",
          "![](", temp_plot, "){width=100%}\n\n",
          
          "## Interpretation\n\n",
          if (results$is_increasing && one_sided_p < 0.05) {
            paste0("There is a statistically significant **INCREASING** trend in the proportion of **'", 
                   first_dep_level, "'** across the ordered levels of **", 
                   ind_var, "** (Z = ", format(z_stat, digits = 3), ", ", format_p_val(one_sided_p), "). ",
                   "As ", ind_var, " increases, the probability of '", first_dep_level, "' significantly increases.")
          } else if (results$is_decreasing && one_sided_p < 0.05) {
            paste0("There is a statistically significant **DECREASING** trend in the proportion of **'", 
                   first_dep_level, "'** across the ordered levels of **", 
                   ind_var, "** (Z = ", format(z_stat, digits = 3), ", ", format_p_val(one_sided_p), "). ",
                   "As ", ind_var, " increases, the probability of '", first_dep_level, "' significantly decreases.")
          } else if (pearson_p < 0.05) {
            paste0("There is a statistically significant relationship between **", 
                   dep_var, "** and **", ind_var, "** (χ² = ", format(pearson_chi_sq, digits = 3), 
                   ", df = ", pearson_df, ", ", format_p_val(pearson_p), "). ",
                   "This suggests a systematic relationship between the variables.")
          } else {
            paste0("No statistically significant trend was detected in the proportion of **",
                   dep_var, "** across the ordered levels of **",
                   ind_var, "** (χ² = ", format(pearson_chi_sq, digits = 3), ", df = ", pearson_df, 
                   ", ", format_p_val(pearson_p), "). ",
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
          "The Pearson's Chi-square test of independence was also performed to assess the overall association between the variables.\n\n",
          
          "---\n",
          "*Report generated using CATrend Analyzer*"
        )
        
        # Write the Rmd content to temporary file
        writeLines(rmd_content, temp_report)
        
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
