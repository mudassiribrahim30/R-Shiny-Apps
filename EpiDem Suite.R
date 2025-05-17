library(shiny)
library(shinydashboard)
library(ggplot2)
library(dplyr)
library(sf)
library(ggpubr)
library(survival)
library(survminer)
library(flextable)
library(officer)
library(DT)
library(haven)
library(readxl)
library(plotly)
library(colourpicker)
library(epitools)

# UI
ui <- dashboardPage(
  dashboardHeader(title = "📈 EpiDem Suite™"),
  dashboardSidebar(
    sidebarMenu(
      menuItem("Data Upload", tabName = "data", icon = icon("database")),
      menuItem("Descriptive Analyses", tabName = "descriptive", icon = icon("chart-bar"),
               menuSubItem("Frequencies", tabName = "freq"),
               menuSubItem("Central Tendency", tabName = "central"),
               menuSubItem("Epidemic Curves", tabName = "epicurve"),
               menuSubItem("Spatial Maps", tabName = "maps"),
               menuSubItem("Demographics", tabName = "demo")),
      menuItem("Disease Frequency", tabName = "disease", icon = icon("viruses"),
               menuSubItem("Incidence/Prevalence", tabName = "incprev"),
               menuSubItem("Attack Rates", tabName = "attack"),
               menuSubItem("Mortality Rates", tabName = "mortality"),
               menuSubItem("Life Table Analysis", tabName = "lifetable")),
      menuItem("Measures of Association", tabName = "association", icon = icon("balance-scale"),
               menuSubItem("Risk Ratios", tabName = "risk"),
               menuSubItem("Odds Ratios", tabName = "odds"),
               menuSubItem("Attributable Risk", tabName = "ar"),
               menuSubItem("Poisson Regression", tabName = "poisson")),
      menuItem("Survival Analysis", tabName = "survival", icon = icon("clock"),
               menuSubItem("Kaplan-Meier", tabName = "km"),
               menuSubItem("Log-Rank Test", tabName = "logrank"),
               menuSubItem("Cox Regression", tabName = "cox")),
      tags$div(
        style = "position: absolute; bottom: 0; width: 100%; padding: 10px; background-color: #222d32; color: white; text-align: center;",
        " Dev: Mudasir Mohammed Ibrahim",
        br(),
        " 📧: mudassiribrahim30@gmail.com"
      )
    )
  ),
  dashboardBody(
    tabItems(
      # Data Upload
      tabItem(tabName = "data",
              fluidRow(
                box(title = "Upload Data", status = "primary", solidHeader = TRUE,
                    fileInput("file1", "Choose Data File",
                              accept = c(".csv", ".xlsx", ".xls", ".sav", ".dta")),
                    conditionalPanel(
                      condition = "input.file1.type == 'text/csv' || input.file1.type == 'text/comma-separated-values,text/plain'",
                      checkboxInput("header", "Header", TRUE),
                      radioButtons("sep", "Separator", choices = c(Comma = ",", Semicolon = ";", Tab = "\t"), selected = ","),
                      radioButtons("quote", "Quote", choices = c(None = "", "Double Quote" = '"', "Single Quote" = "'"), selected = '"')
                    ),
                    conditionalPanel(
                      condition = "input.file1.type == 'application/vnd.ms-excel' || input.file1.type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'",
                      numericInput("sheet", "Sheet Number", value = 1)
                    )
                ),
                box(title = "Data Preview", status = "info", solidHeader = TRUE,
                    DTOutput("contents")
                )
              )
      ),
      
      # Frequencies
      tabItem(tabName = "freq",
              fluidRow(
                box(title = "Frequency Analysis", status = "primary", solidHeader = TRUE,
                    selectInput("freq_var", "Select Variable:", choices = NULL),
                    radioButtons("freq_type", "Output Type:", choices = c("Table", "Plot"), selected = "Table"),
                    conditionalPanel(
                      condition = "input.freq_type == 'Plot'",
                      colourInput("freq_color", "Select Color:", value = "steelblue"),
                      textInput("freq_title", "Plot Title:", value = "Frequency Plot"),
                      numericInput("freq_title_size", "Title Size:", value = 14),
                      numericInput("freq_x_size", "X-axis Label Size:", value = 12),
                      numericInput("freq_y_size", "Y-axis Label Size:", value = 12)
                    )
                ),
                box(title = "Results", status = "info", solidHeader = TRUE,
                    conditionalPanel(
                      condition = "input.freq_type == 'Table'",
                      DTOutput("freq_table")
                    ),
                    conditionalPanel(
                      condition = "input.freq_type == 'Plot'",
                      plotlyOutput("freq_plot")
                    ),
                    downloadButton("download_freq", "Download Results")
                )
              )
      ),
      
      # Central Tendency
      tabItem(tabName = "central",
              fluidRow(
                box(title = "Central Tendency", status = "primary", solidHeader = TRUE,
                    selectInput("central_var", "Select Numeric Variable:", choices = NULL),
                    radioButtons("central_type", "Output Type:", 
                                 choices = c("Table", "Boxplot", "Histogram"), 
                                 selected = "Table"),
                    conditionalPanel(
                      condition = "input.central_type != 'Table'",
                      colourInput("central_color", "Select Color:", value = "steelblue"),
                      textInput("central_title", "Plot Title:", value = "Distribution Plot"),
                      numericInput("central_title_size", "Title Size:", value = 14),
                      numericInput("central_x_size", "X-axis Label Size:", value = 12),
                      numericInput("central_y_size", "Y-axis Label Size:", value = 12),
                      conditionalPanel(
                        condition = "input.central_type == 'Histogram'",
                        numericInput("central_bins", "Number of Bins:", value = 30)
                      )
                    )
                ),
                box(title = "Results", status = "info", solidHeader = TRUE,
                    conditionalPanel(
                      condition = "input.central_type == 'Table'",
                      DTOutput("central_table")
                    ),
                    conditionalPanel(
                      condition = "input.central_type == 'Boxplot'",
                      plotlyOutput("central_boxplot")
                    ),
                    conditionalPanel(
                      condition = "input.central_type == 'Histogram'",
                      plotlyOutput("central_histogram")
                    ),
                    downloadButton("download_central", "Download Results")
                )
              )
      ),
      
      # Epidemic Curve
      tabItem(tabName = "epicurve",
              fluidRow(
                box(title = "Epidemic Curve", status = "primary", solidHeader = TRUE,
                    selectInput("epi_date", "Select Date Variable:", choices = NULL),
                    selectInput("epi_case", "Select Case Variable:", choices = NULL),
                    selectInput("epi_group", "Group By (optional):", choices = NULL),
                    radioButtons("epi_interval", "Time Interval:", 
                                 choices = c("Day", "Week", "Month", "Year"), selected = "Week"),
                    colourInput("epi_color", "Select Color:", value = "steelblue"),
                    textInput("epi_title", "Plot Title:", value = "Epidemic Curve"),
                    numericInput("epi_title_size", "Title Size:", value = 14),
                    numericInput("epi_x_size", "X-axis Label Size:", value = 12),
                    numericInput("epi_y_size", "Y-axis Label Size:", value = 12)
                ),
                box(title = "Epidemic Curve", status = "info", solidHeader = TRUE,
                    plotlyOutput("epi_curve"),
                    downloadButton("download_epicurve", "Download Plot")
                )
              )
      ),
      
      # Spatial Maps
      tabItem(tabName = "maps",
              fluidRow(
                box(title = "Spatial Distribution", status = "primary", solidHeader = TRUE,
                    selectInput("map_var", "Select Variable to Map:", choices = NULL),
                    selectInput("map_region", "Select Region Variable:", choices = NULL),
                    radioButtons("map_type", "Map Type:", choices = c("Choropleth", "Point"), selected = "Choropleth"),
                    colourInput("map_color", "Select Color:", value = "steelblue"),
                    textInput("map_title", "Plot Title:", value = "Spatial Distribution"),
                    numericInput("map_title_size", "Title Size:", value = 14),
                    numericInput("map_x_size", "X-axis Label Size:", value = 12),
                    numericInput("map_y_size", "Y-axis Label Size:", value = 12)
                ),
                box(title = "Map", status = "info", solidHeader = TRUE,
                    plotlyOutput("map_plot"),
                    downloadButton("download_map", "Download Map")
                )
              )
      ),
      
      # Demographics
      tabItem(tabName = "demo",
              fluidRow(
                box(title = "Demographic Breakdown", status = "primary", solidHeader = TRUE,
                    selectInput("demo_var", "Select Outcome Variable:", choices = NULL),
                    selectInput("demo_group", "Select Grouping Variable(s):", choices = NULL, multiple = TRUE)
                ),
                box(title = "Results", status = "info", solidHeader = TRUE,
                    div(style = 'overflow-x: scroll', DTOutput("demo_table")),
                    downloadButton("download_demo", "Download Table")
                )
              )
      ),
      
      # Incidence/Prevalence
      tabItem(tabName = "incprev",
              fluidRow(
                box(title = "Incidence/Prevalence", status = "primary", solidHeader = TRUE,
                    selectInput("ip_case", "Select Case Variable:", choices = NULL),
                    selectInput("ip_pop", "Select Population Variable:", choices = NULL),
                    selectInput("ip_time", "Select Time Variable:", choices = NULL),
                    radioButtons("ip_measure", "Measure:", 
                                 choices = c("Incidence Rate", "Cumulative Incidence", "Prevalence"), 
                                 selected = "Incidence Rate")
                ),
                box(title = "Results", status = "info", solidHeader = TRUE,
                    DTOutput("ip_table"),
                    downloadButton("download_ip", "Download Table")
                )
              )
      ),
      
      # Attack Rates
      tabItem(tabName = "attack",
              fluidRow(
                box(title = "Attack Rate", status = "primary", solidHeader = TRUE,
                    selectInput("ar_case", "Select Case Variable:", choices = NULL),
                    selectInput("ar_pop", "Select Population Variable:", choices = NULL),
                    selectInput("ar_group", "Group By (optional):", choices = NULL),
                    radioButtons("ar_format", "Output Format:", 
                                 choices = c("Table", "Detailed Table"), 
                                 selected = "Table")
                ),
                box(title = "Results", status = "info", solidHeader = TRUE,
                    conditionalPanel(
                      condition = "input.ar_format == 'Table'",
                      DTOutput("ar_table")
                    ),
                    conditionalPanel(
                      condition = "input.ar_format == 'Detailed Table'",
                      DTOutput("ar_detailed_table")
                    ),
                    downloadButton("download_ar", "Download Table")
                )
              )
      ),
      
      # Mortality Rates
      tabItem(tabName = "mortality",
              fluidRow(
                box(title = "Mortality Rates", status = "primary", solidHeader = TRUE,
                    selectInput("mort_case", "Select Death Variable:", choices = NULL),
                    selectInput("mort_pop", "Select Population Variable:", choices = NULL),
                    radioButtons("mort_measure", "Measure:", 
                                 choices = c("Mortality Rate", "Case Fatality Rate"), 
                                 selected = "Mortality Rate")
                ),
                box(title = "Results", status = "info", solidHeader = TRUE,
                    DTOutput("mort_table"),
                    downloadButton("download_mort", "Download Table")
                )
              )
      ),
      
      # Life Table Analysis
      tabItem(tabName = "lifetable",
              fluidRow(
                box(title = "Life Table Analysis", status = "primary", solidHeader = TRUE,
                    selectInput("lt_time", "Select Time Variable:", choices = NULL),
                    selectInput("lt_event", "Select Event Variable:", choices = NULL),
                    textInput("lt_breaks", "Time Breaks (comma separated):", value = "0,10,20,30,40,50,60,70,80"),
                    textInput("lt_title", "Table Title:", value = "Life Table Analysis")
                ),
                box(title = "Life Table Results", status = "info", solidHeader = TRUE,
                    div(style = 'overflow-x: scroll', DTOutput("lifetable_results")),
                    downloadButton("download_lifetable", "Download Table")
                )
              )
      ),
      
      # Risk Ratios
      tabItem(tabName = "risk",
              fluidRow(
                box(title = "Risk Ratio", status = "primary", solidHeader = TRUE,
                    selectInput("rr_outcome", "Select Outcome Variable:", choices = NULL),
                    selectInput("rr_exposure", "Select Exposure Variable:", choices = NULL),
                    numericInput("rr_conf", "Confidence Level:", value = 0.95, min = 0.5, max = 0.99, step = 0.01),
                    textInput("rr_exposed_label", "Exposed Group Label:", value = "Exposed"),
                    textInput("rr_unexposed_label", "Unexposed Group Label:", value = "Unexposed")
                ),
                box(title = "Results", status = "info", solidHeader = TRUE,
                    verbatimTextOutput("rr_results"),
                    DTOutput("rr_table"),
                    downloadButton("download_rr", "Download Table")
                )
              )
      ),
      
      # Odds Ratios
      tabItem(tabName = "odds",
              fluidRow(
                box(title = "Odds Ratio", status = "primary", solidHeader = TRUE,
                    selectInput("or_outcome", "Select Outcome Variable:", choices = NULL),
                    selectInput("or_exposure", "Select Exposure Variable:", choices = NULL),
                    numericInput("or_conf", "Confidence Level:", value = 0.95, min = 0.5, max = 0.99, step = 0.01),
                    textInput("or_exposed_label", "Exposed Group Label:", value = "Exposed"),
                    textInput("or_unexposed_label", "Unexposed Group Label:", value = "Unexposed")
                ),
                box(title = "Results", status = "info", solidHeader = TRUE,
                    verbatimTextOutput("or_results"),
                    DTOutput("or_table"),
                    downloadButton("download_or", "Download Table")
                )
              )
      ),
      
      # Attributable Risk
      tabItem(tabName = "ar",
              fluidRow(
                box(title = "Attributable Risk", status = "primary", solidHeader = TRUE,
                    selectInput("ar_outcome", "Select Outcome Variable:", choices = NULL),
                    selectInput("ar_exposure", "Select Exposure Variable:", choices = NULL),
                    numericInput("ar_conf", "Confidence Level:", value = 0.95, min = 0.5, max = 0.99, step = 0.01),
                    checkboxInput("ar_paf", "Calculate Population Attributable Fraction", value = FALSE),
                    textInput("ar_exposed_label", "Exposed Group Label:", value = "Exposed"),
                    textInput("ar_unexposed_label", "Unexposed Group Label:", value = "Unexposed")
                ),
                box(title = "Results", status = "info", solidHeader = TRUE,
                    verbatimTextOutput("ar_results"),
                    DTOutput("ar_table"),
                    downloadButton("download_arisk", "Download Table")
                )
              )
      ),
      
      # Poisson Regression
      tabItem(tabName = "poisson",
              fluidRow(
                box(title = "Poisson Regression", status = "primary", solidHeader = TRUE,
                    selectInput("poisson_outcome", "Select Outcome Count Variable:", choices = NULL),
                    selectInput("poisson_offset", "Select Offset Variable (log transform will be applied):", choices = NULL),
                    selectInput("poisson_vars", "Select Predictor Variables:", choices = NULL, multiple = TRUE),
                    uiOutput("poisson_ref_ui"),
                    numericInput("poisson_conf", "Confidence Level:", value = 0.95, min = 0.5, max = 0.99, step = 0.01),
                    checkboxInput("poisson_irr", "Show Incidence Rate Ratios (IRR)", value = TRUE)
                ),
                box(title = "Poisson Regression Results", status = "info", solidHeader = TRUE,
                    verbatimTextOutput("poisson_summary"),
                    DTOutput("poisson_table"),
                    downloadButton("download_poisson", "Download Results")
                )
              )
      ),
      
      # Kaplan-Meier
      tabItem(tabName = "km",
              fluidRow(
                box(title = "Kaplan-Meier Analysis", status = "primary", solidHeader = TRUE,
                    selectInput("km_time", "Select Time Variable:", choices = NULL),
                    selectInput("km_event", "Select Event Variable:", choices = NULL),
                    textInput("km_time_entry", "Manual Time Entry (comma separated):", placeholder = "e.g., 10,20,30"),
                    textInput("km_event_entry", "Manual Event Entry (comma separated 0/1):", placeholder = "e.g., 1,0,1"),
                    selectInput("km_group", "Group By (optional):", choices = NULL),
                    numericInput("km_conf", "Confidence Level:", value = 0.95, min = 0.5, max = 0.99, step = 0.01),
                    checkboxInput("km_risktable", "Show Risk Table", value = TRUE),
                    colourInput("km_color1", "Color for Group 1:", value = "#E41A1C"),
                    colourInput("km_color2", "Color for Group 2:", value = "#377EB8"),
                    textInput("km_title", "Plot Title:", value = "Kaplan-Meier Survival Curve"),
                    numericInput("km_title_size", "Title Size:", value = 14),
                    numericInput("km_x_size", "X-axis Label Size:", value = 12),
                    numericInput("km_y_size", "Y-axis Label Size:", value = 12)
                ),
                box(title = "Kaplan-Meier Plot", status = "info", solidHeader = TRUE,
                    plotOutput("km_plot"),
                    downloadButton("download_km", "Download Plot")
                )
              )
      ),
      
      # Log-Rank Test
      tabItem(tabName = "logrank",
              fluidRow(
                box(title = "Log-Rank Test", status = "primary", solidHeader = TRUE,
                    selectInput("lr_time", "Select Time Variable:", choices = NULL),
                    selectInput("lr_event", "Select Event Variable:", choices = NULL),
                    selectInput("lr_group", "Select Grouping Variable:", choices = NULL)
                ),
                box(title = "Results", status = "info", solidHeader = TRUE,
                    verbatimTextOutput("logrank_results"),
                    downloadButton("download_logrank", "Download Results")
                )
              )
      ),
      
      # Cox Regression
      tabItem(tabName = "cox",
              fluidRow(
                box(title = "Cox Regression", status = "primary", solidHeader = TRUE,
                    selectInput("cox_time", "Select Time Variable:", choices = NULL),
                    selectInput("cox_event", "Select Event Variable:", choices = NULL),
                    selectInput("cox_vars", "Select Predictor Variables:", choices = NULL, multiple = TRUE),
                    numericInput("cox_conf", "Confidence Level:", value = 0.95, min = 0.5, max = 0.99, step = 0.01)
                ),
                box(title = "Results", status = "info", solidHeader = TRUE,
                    verbatimTextOutput("cox_summary"),
                    downloadButton("download_cox", "Download Results")
                )
              )
      )
    )
  )
)

# Server function
server <- function(input, output, session) {
  
  # Reactive data
  data <- reactive({
    req(input$file1)
    
    ext <- tools::file_ext(input$file1$name)
    
    df <- switch(ext,
                 csv = read.csv(input$file1$datapath,
                                header = input$header,
                                sep = input$sep,
                                quote = input$quote),
                 xls = read_excel(input$file1$datapath, sheet = input$sheet),
                 xlsx = read_excel(input$file1$datapath, sheet = input$sheet),
                 sav = read_sav(input$file1$datapath),
                 dta = read_dta(input$file1$datapath),
                 validate("Invalid file; Please upload a .csv, .xls, .xlsx, .sav, or .dta file")
    )
    
    # Convert to data.frame if it's a tibble
    if (inherits(df, "tbl_df")) {
      df <- as.data.frame(df)
    }
    
    # Update all select inputs
    updateSelectInput(session, "freq_var", choices = names(df))
    updateSelectInput(session, "central_var", choices = names(df)[sapply(df, is.numeric)])
    updateSelectInput(session, "epi_date", choices = names(df)[sapply(df, function(x) inherits(x, "Date") | is.character(x))])
    updateSelectInput(session, "epi_case", choices = names(df))
    updateSelectInput(session, "epi_group", choices = c("None", names(df)))
    updateSelectInput(session, "map_var", choices = names(df))
    updateSelectInput(session, "map_region", choices = names(df))
    updateSelectInput(session, "demo_var", choices = names(df))
    updateSelectInput(session, "demo_group", choices = names(df))
    updateSelectInput(session, "ip_case", choices = names(df))
    updateSelectInput(session, "ip_pop", choices = names(df)[sapply(df, is.numeric)])
    updateSelectInput(session, "ip_time", choices = names(df))
    updateSelectInput(session, "ar_case", choices = names(df))
    updateSelectInput(session, "ar_pop", choices = names(df)[sapply(df, is.numeric)])
    updateSelectInput(session, "ar_group", choices = c("None", names(df)))
    updateSelectInput(session, "mort_case", choices = names(df))
    updateSelectInput(session, "mort_pop", choices = names(df)[sapply(df, is.numeric)])
    updateSelectInput(session, "lt_time", choices = names(df)[sapply(df, is.numeric)])
    updateSelectInput(session, "lt_event", choices = names(df))
    updateSelectInput(session, "rr_outcome", choices = names(df))
    updateSelectInput(session, "rr_exposure", choices = names(df))
    updateSelectInput(session, "or_outcome", choices = names(df))
    updateSelectInput(session, "or_exposure", choices = names(df))
    updateSelectInput(session, "ar_outcome", choices = names(df))
    updateSelectInput(session, "ar_exposure", choices = names(df))
    updateSelectInput(session, "poisson_outcome", choices = names(df)[sapply(df, is.numeric)])
    updateSelectInput(session, "poisson_offset", choices = names(df)[sapply(df, is.numeric)])
    updateSelectInput(session, "poisson_vars", choices = names(df))
    updateSelectInput(session, "km_time", choices = names(df)[sapply(df, is.numeric)])
    updateSelectInput(session, "km_event", choices = names(df))
    updateSelectInput(session, "km_group", choices = c("None", names(df)))
    updateSelectInput(session, "lr_time", choices = names(df)[sapply(df, is.numeric)])
    updateSelectInput(session, "lr_event", choices = names(df))
    updateSelectInput(session, "lr_group", choices = names(df))
    updateSelectInput(session, "cox_time", choices = names(df)[sapply(df, is.numeric)])
    updateSelectInput(session, "cox_event", choices = names(df))
    updateSelectInput(session, "cox_vars", choices = names(df))
    
    df
  })
  
  # Data preview
  output$contents <- renderDT({
    datatable(data(), options = list(scrollX = TRUE))
  })
  
  # Frequency analysis
  output$freq_table <- renderDT({
    req(input$freq_var)
    df <- data()
    freq_table <- as.data.frame(table(df[[input$freq_var]]))
    names(freq_table) <- c(input$freq_var, "Frequency")
    freq_table$Proportion <- freq_table$Frequency / sum(freq_table$Frequency)
    datatable(freq_table)
  })
  
  output$freq_plot <- renderPlotly({
    req(input$freq_var)
    df <- data()
    
    p <- ggplot(df, aes(x = .data[[input$freq_var]])) +
      geom_bar(fill = input$freq_color) +
      labs(title = input$freq_title, x = input$freq_var, y = "Count") +
      theme_minimal() +
      theme(
        plot.title = element_text(size = input$freq_title_size),
        axis.title.x = element_text(size = input$freq_x_size),
        axis.title.y = element_text(size = input$freq_y_size)
      )
    
    ggplotly(p)
  })
  
  output$download_freq <- downloadHandler(
    filename = function() {
      if(input$freq_type == "Table") {
        "frequency_table.docx"
      } else {
        "frequency_plot.png"
      }
    },
    content = function(file) {
      if(input$freq_type == "Table") {
        df <- data()
        freq_table <- as.data.frame(table(df[[input$freq_var]]))
        names(freq_table) <- c(input$freq_var, "Frequency")
        freq_table$Proportion <- freq_table$Frequency / sum(freq_table$Frequency)
        
        ft <- flextable(freq_table)
        doc <- read_docx() %>% 
          body_add_flextable(ft)
        print(doc, target = file)
      } else {
        df <- data()
        p <- ggplot(df, aes(x = .data[[input$freq_var]])) +
          geom_bar(fill = input$freq_color) +
          labs(title = input$freq_title, x = input$freq_var, y = "Count") +
          theme_minimal() +
          theme(
            plot.title = element_text(size = input$freq_title_size),
            axis.title.x = element_text(size = input$freq_x_size),
            axis.title.y = element_text(size = input$freq_y_size)
          )
        ggsave(file, plot = p, device = "png", width = 8, height = 6)
      }
    }
  )
  
  # Central tendency
  output$central_table <- renderDT({
    req(input$central_var)
    df <- data()
    var <- df[[input$central_var]]
    
    # Calculate mode
    get_mode <- function(x) {
      ux <- unique(x)
      ux[which.max(tabulate(match(x, ux)))]
    }
    
    central_table <- data.frame(
      Measure = c("Mean", "Median", "SD", "Min", "Max", "IQR", "Mode"),
      Value = c(mean(var, na.rm = TRUE), 
                median(var, na.rm = TRUE),
                sd(var, na.rm = TRUE),
                min(var, na.rm = TRUE),
                max(var, na.rm = TRUE),
                IQR(var, na.rm = TRUE),
                get_mode(var))
    )
    datatable(central_table)
  })
  
  output$central_boxplot <- renderPlotly({
    req(input$central_var)
    df <- data()
    
    p <- ggplot(df, aes(y = .data[[input$central_var]])) +
      geom_boxplot(fill = input$central_color) +
      labs(title = input$central_title, y = input$central_var) +
      theme_minimal() +
      theme(
        plot.title = element_text(size = input$central_title_size),
        axis.title.x = element_text(size = input$central_x_size),
        axis.title.y = element_text(size = input$central_y_size)
      )
    
    ggplotly(p)
  })
  
  output$central_histogram <- renderPlotly({
    req(input$central_var)
    df <- data()
    
    p <- ggplot(df, aes(x = .data[[input$central_var]])) +
      geom_histogram(fill = input$central_color, bins = input$central_bins) +
      labs(title = input$central_title, x = input$central_var, y = "Count") +
      theme_minimal() +
      theme(
        plot.title = element_text(size = input$central_title_size),
        axis.title.x = element_text(size = input$central_x_size),
        axis.title.y = element_text(size = input$central_y_size)
      )
    
    ggplotly(p)
  })
  
  output$download_central <- downloadHandler(
    filename = function() {
      if(input$central_type == "Table") {
        "central_tendency_table.docx"
      } else if(input$central_type == "Boxplot") {
        "central_tendency_boxplot.png"
      } else {
        "central_tendency_histogram.png"
      }
    },
    content = function(file) {
      df <- data()
      var <- df[[input$central_var]]
      
      if(input$central_type == "Table") {
        get_mode <- function(x) {
          ux <- unique(x)
          ux[which.max(tabulate(match(x, ux)))]
        }
        
        central_table <- data.frame(
          Measure = c("Mean", "Median", "SD", "Min", "Max", "IQR", "Mode"),
          Value = c(mean(var, na.rm = TRUE), 
                    median(var, na.rm = TRUE),
                    sd(var, na.rm = TRUE),
                    min(var, na.rm = TRUE),
                    max(var, na.rm = TRUE),
                    IQR(var, na.rm = TRUE),
                    get_mode(var))
        )
        
        ft <- flextable(central_table)
        doc <- read_docx() %>% 
          body_add_flextable(ft)
        print(doc, target = file)
      } else if(input$central_type == "Boxplot") {
        p <- ggplot(df, aes(y = .data[[input$central_var]])) +
          geom_boxplot(fill = input$central_color) +
          labs(title = input$central_title, y = input$central_var) +
          theme_minimal() +
          theme(
            plot.title = element_text(size = input$central_title_size),
            axis.title.x = element_text(size = input$central_x_size),
            axis.title.y = element_text(size = input$central_y_size)
          )
        ggsave(file, plot = p, device = "png", width = 8, height = 6)
      } else {
        p <- ggplot(df, aes(x = .data[[input$central_var]])) +
          geom_histogram(fill = input$central_color, bins = input$central_bins) +
          labs(title = input$central_title, x = input$central_var, y = "Count") +
          theme_minimal() +
          theme(
            plot.title = element_text(size = input$central_title_size),
            axis.title.x = element_text(size = input$central_x_size),
            axis.title.y = element_text(size = input$central_y_size)
          )
        ggsave(file, plot = p, device = "png", width = 8, height = 6)
      }
    }
  )
  
  # Epidemic Curve
  output$epi_curve <- renderPlotly({
    req(input$epi_date, input$epi_case)
    df <- data()
    
    # Convert date if needed
    if(is.character(df[[input$epi_date]])) {
      df[[input$epi_date]] <- as.Date(df[[input$epi_date]])
    }
    
    # Aggregate data based on time interval
    if(input$epi_interval == "Day") {
      df$time_group <- as.Date(df[[input$epi_date]])
    } else if(input$epi_interval == "Week") {
      df$time_group <- cut(df[[input$epi_date]], "week")
    } else if(input$epi_interval == "Month") {
      df$time_group <- cut(df[[input$epi_date]], "month")
    } else {
      df$time_group <- cut(df[[input$epi_date]], "year")
    }
    
    # Create plot
    if(input$epi_group == "None") {
      epi_data <- df %>%
        group_by(time_group) %>%
        summarise(cases = sum(.data[[input$epi_case]], na.rm = TRUE))
      
      p <- ggplot(epi_data, aes(x = time_group, y = cases)) +
        geom_col(fill = input$epi_color) +
        labs(title = input$epi_title, x = "Time", y = "Number of Cases") +
        theme_minimal() +
        theme(
          plot.title = element_text(size = input$epi_title_size),
          axis.title.x = element_text(size = input$epi_x_size),
          axis.title.y = element_text(size = input$epi_y_size),
          axis.text.x = element_text(angle = 45, hjust = 1)
        )
    } else {
      epi_data <- df %>%
        group_by(time_group, .data[[input$epi_group]]) %>%
        summarise(cases = sum(.data[[input$epi_case]], na.rm = TRUE))
      
      p <- ggplot(epi_data, aes(x = time_group, y = cases, fill = .data[[input$epi_group]])) +
        geom_col(position = "dodge") +
        labs(title = input$epi_title, x = "Time", y = "Number of Cases", fill = input$epi_group) +
        theme_minimal() +
        theme(
          plot.title = element_text(size = input$epi_title_size),
          axis.title.x = element_text(size = input$epi_x_size),
          axis.title.y = element_text(size = input$epi_y_size),
          axis.text.x = element_text(angle = 45, hjust = 1)
        )
    }
    
    ggplotly(p)
  })
  
  output$download_epicurve <- downloadHandler(
    filename = "epidemic_curve.png",
    content = function(file) {
      df <- data()
      
      if(is.character(df[[input$epi_date]])) {
        df[[input$epi_date]] <- as.Date(df[[input$epi_date]])
      }
      
      if(input$epi_interval == "Day") {
        df$time_group <- as.Date(df[[input$epi_date]])
      } else if(input$epi_interval == "Week") {
        df$time_group <- cut(df[[input$epi_date]], "week")
      } else if(input$epi_interval == "Month") {
        df$time_group <- cut(df[[input$epi_date]], "month")
      } else {
        df$time_group <- cut(df[[input$epi_date]], "year")
      }
      
      if(input$epi_group == "None") {
        epi_data <- df %>%
          group_by(time_group) %>%
          summarise(cases = sum(.data[[input$epi_case]], na.rm = TRUE))
        
        p <- ggplot(epi_data, aes(x = time_group, y = cases)) +
          geom_col(fill = input$epi_color) +
          labs(title = input$epi_title, x = "Time", y = "Number of Cases") +
          theme_minimal() +
          theme(
            plot.title = element_text(size = input$epi_title_size),
            axis.title.x = element_text(size = input$epi_x_size),
            axis.title.y = element_text(size = input$epi_y_size),
            axis.text.x = element_text(angle = 45, hjust = 1)
          )
      } else {
        epi_data <- df %>%
          group_by(time_group, .data[[input$epi_group]]) %>%
          summarise(cases = sum(.data[[input$epi_case]], na.rm = TRUE))
        
        p <- ggplot(epi_data, aes(x = time_group, y = cases, fill = .data[[input$epi_group]])) +
          geom_col(position = "dodge") +
          labs(title = input$epi_title, x = "Time", y = "Number of Cases", fill = input$epi_group) +
          theme_minimal() +
          theme(
            plot.title = element_text(size = input$epi_title_size),
            axis.title.x = element_text(size = input$epi_x_size),
            axis.title.y = element_text(size = input$epi_y_size),
            axis.text.x = element_text(angle = 45, hjust = 1)
          )
      }
      
      ggsave(file, plot = p, device = "png", width = 10, height = 6)
    }
  )
  
  # Spatial maps (simplified - would need spatial data for full implementation)
  output$map_plot <- renderPlotly({
    req(input$map_var, input$map_region)
    df <- data()
    
    if(input$map_type == "Choropleth") {
      # Simplified choropleth - would need spatial data for real implementation
      region_counts <- df %>%
        group_by(.data[[input$map_region]]) %>%
        summarise(value = mean(.data[[input$map_var]], na.rm = TRUE))
      
      p <- ggplot(region_counts, aes(x = .data[[input$map_region]], y = value, fill = value)) +
        geom_col() +
        scale_fill_gradient(low = "blue", high = "red") +
        labs(title = input$map_title, 
             x = input$map_region, y = input$map_var) +
        theme_minimal() +
        theme(
          plot.title = element_text(size = input$map_title_size),
          axis.title.x = element_text(size = input$map_x_size),
          axis.title.y = element_text(size = input$map_y_size),
          axis.text.x = element_text(angle = 45, hjust = 1)
        )
    } else {
      # Point map - simplified
      p <- ggplot(df, aes(x = .data[[input$map_region]], y = .data[[input$map_var]])) +
        geom_point(color = input$map_color, size = 3) +
        labs(title = input$map_title, 
             x = input$map_region, y = input$map_var) +
        theme_minimal() +
        theme(
          plot.title = element_text(size = input$map_title_size),
          axis.title.x = element_text(size = input$map_x_size),
          axis.title.y = element_text(size = input$map_y_size)
        )
    }
    
    ggplotly(p)
  })
  
  output$download_map <- downloadHandler(
    filename = "spatial_map.png",
    content = function(file) {
      df <- data()
      
      if(input$map_type == "Choropleth") {
        region_counts <- df %>%
          group_by(.data[[input$map_region]]) %>%
          summarise(value = mean(.data[[input$map_var]], na.rm = TRUE))
        
        p <- ggplot(region_counts, aes(x = .data[[input$map_region]], y = value, fill = value)) +
          geom_col() +
          scale_fill_gradient(low = "blue", high = "red") +
          labs(title = input$map_title, 
               x = input$map_region, y = input$map_var) +
          theme_minimal() +
          theme(
            plot.title = element_text(size = input$map_title_size),
            axis.title.x = element_text(size = input$map_x_size),
            axis.title.y = element_text(size = input$map_y_size),
            axis.text.x = element_text(angle = 45, hjust = 1)
          )
      } else {
        p <- ggplot(df, aes(x = .data[[input$map_region]], y = .data[[input$map_var]])) +
          geom_point(color = input$map_color, size = 3) +
          labs(title = input$map_title, 
               x = input$map_region, y = input$map_var) +
          theme_minimal() +
          theme(
            plot.title = element_text(size = input$map_title_size),
            axis.title.x = element_text(size = input$map_x_size),
            axis.title.y = element_text(size = input$map_y_size)
          )
      }
      
      ggsave(file, plot = p, device = "png", width = 10, height = 6)
    }
  )
  
  # Demographic breakdown
  output$demo_table <- renderDT({
    req(input$demo_var, input$demo_group)
    df <- data()
    
    if(length(input$demo_group) == 1) {
      demo_table <- df %>%
        group_by(.data[[input$demo_group]]) %>%
        summarise(
          Count = n(),
          Mean = mean(.data[[input$demo_var]], na.rm = TRUE),
          SD = sd(.data[[input$demo_var]], na.rm = TRUE),
          Median = median(.data[[input$demo_var]], na.rm = TRUE),
          Min = min(.data[[input$demo_var]], na.rm = TRUE),
          Max = max(.data[[input$demo_var]], na.rm = TRUE),
          Mode = {
            ux <- unique(.data[[input$demo_var]])
            ux[which.max(tabulate(match(.data[[input$demo_var]], ux)))]
          }
        )
    } else {
      demo_table <- df %>%
        group_by(across(all_of(input$demo_group))) %>%
        summarise(
          Count = n(),
          Mean = mean(.data[[input$demo_var]], na.rm = TRUE),
          SD = sd(.data[[input$demo_var]], na.rm = TRUE),
          Median = median(.data[[input$demo_var]], na.rm = TRUE),
          Min = min(.data[[input$demo_var]], na.rm = TRUE),
          Max = max(.data[[input$demo_var]], na.rm = TRUE),
          Mode = {
            ux <- unique(.data[[input$demo_var]])
            ux[which.max(tabulate(match(.data[[input$demo_var]], ux)))]
          }
        )
    }
    
    datatable(demo_table)
  })
  
  output$download_demo <- downloadHandler(
    filename = "demographic_table.docx",
    content = function(file) {
      df <- data()
      
      if(length(input$demo_group) == 1) {
        demo_table <- df %>%
          group_by(.data[[input$demo_group]]) %>%
          summarise(
            Count = n(),
            Mean = mean(.data[[input$demo_var]], na.rm = TRUE),
            SD = sd(.data[[input$demo_var]], na.rm = TRUE),
            Median = median(.data[[input$demo_var]], na.rm = TRUE),
            Min = min(.data[[input$demo_var]], na.rm = TRUE),
            Max = max(.data[[input$demo_var]], na.rm = TRUE),
            Mode = {
              ux <- unique(.data[[input$demo_var]])
              ux[which.max(tabulate(match(.data[[input$demo_var]], ux)))]
            }
          )
      } else {
        demo_table <- df %>%
          group_by(across(all_of(input$demo_group))) %>%
          summarise(
            Count = n(),
            Mean = mean(.data[[input$demo_var]], na.rm = TRUE),
            SD = sd(.data[[input$demo_var]], na.rm = TRUE),
            Median = median(.data[[input$demo_var]], na.rm = TRUE),
            Min = min(.data[[input$demo_var]], na.rm = TRUE),
            Max = max(.data[[input$demo_var]], na.rm = TRUE),
            Mode = {
              ux <- unique(.data[[input$demo_var]])
              ux[which.max(tabulate(match(.data[[input$demo_var]], ux)))]
            }
          )
      }
      
      ft <- flextable(demo_table)
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Incidence/Prevalence
  output$ip_table <- renderDT({
    req(input$ip_case, input$ip_pop, input$ip_time)
    df <- data()
    
    if(input$ip_measure == "Incidence Rate") {
      ip_data <- df %>%
        summarise(Cases = sum(.data[[input$ip_case]], na.rm = TRUE),
                  Population = sum(.data[[input$ip_pop]], na.rm = TRUE),
                  Time = n_distinct(.data[[input$ip_time]], na.rm = TRUE),
                  Incidence_Rate = (Cases / Population) / Time * 1000)
    } else if(input$ip_measure == "Cumulative Incidence") {
      ip_data <- df %>%
        summarise(Cases = sum(.data[[input$ip_case]], na.rm = TRUE),
                  Population = sum(.data[[input$ip_pop]], na.rm = TRUE),
                  Cumulative_Incidence = Cases / Population * 1000)
    } else {
      ip_data <- df %>%
        summarise(Cases = sum(.data[[input$ip_case]], na.rm = TRUE),
                  Population = sum(.data[[input$ip_pop]], na.rm = TRUE),
                  Prevalence = Cases / Population * 1000)
    }
    
    datatable(ip_data)
  })
  
  output$download_ip <- downloadHandler(
    filename = "incidence_prevalence_table.docx",
    content = function(file) {
      df <- data()
      
      if(input$ip_measure == "Incidence Rate") {
        ip_data <- df %>%
          summarise(Cases = sum(.data[[input$ip_case]], na.rm = TRUE),
                    Population = sum(.data[[input$ip_pop]], na.rm = TRUE),
                    Time = n_distinct(.data[[input$ip_time]], na.rm = TRUE),
                    Incidence_Rate = (Cases / Population) / Time * 1000)
      } else if(input$ip_measure == "Cumulative Incidence") {
        ip_data <- df %>%
          summarise(Cases = sum(.data[[input$ip_case]], na.rm = TRUE),
                    Population = sum(.data[[input$ip_pop]], na.rm = TRUE),
                    Cumulative_Incidence = Cases / Population * 1000)
      } else {
        ip_data <- df %>%
          summarise(Cases = sum(.data[[input$ip_case]], na.rm = TRUE),
                    Population = sum(.data[[input$ip_pop]], na.rm = TRUE),
                    Prevalence = Cases / Population * 1000)
      }
      
      ft <- flextable(ip_data)
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Attack rate - enhanced with detailed table option
  output$ar_table <- renderDT({
    req(input$ar_case, input$ar_pop)
    df <- data()
    
    if(input$ar_group == "None") {
      ar_data <- df %>%
        summarise(Cases = sum(.data[[input$ar_case]], na.rm = TRUE),
                  Population = sum(.data[[input$ar_pop]], na.rm = TRUE),
                  Attack_Rate = Cases / Population * 100)
    } else {
      ar_data <- df %>%
        group_by(.data[[input$ar_group]]) %>%
        summarise(Cases = sum(.data[[input$ar_case]], na.rm = TRUE),
                  Population = sum(.data[[input$ar_pop]], na.rm = TRUE),
                  Attack_Rate = Cases / Population * 100)
    }
    
    datatable(ar_data, options = list(pageLength = 10))
  })
  
  output$ar_detailed_table <- renderDT({
    req(input$ar_case, input$ar_pop)
    df <- data()
    
    if(input$ar_group == "None") {
      ar_data <- df %>%
        summarise(Cases = sum(.data[[input$ar_case]], na.rm = TRUE),
                  Population = sum(.data[[input$ar_pop]], na.rm = TRUE),
                  Attack_Rate = Cases / Population * 100,
                  Lower_CI = prop.test(Cases, Population, conf.level = input$ar_conf)$conf.int[1] * 100,
                  Upper_CI = prop.test(Cases, Population, conf.level = input$ar_conf)$conf.int[2] * 100)
    } else {
      ar_data <- df %>%
        group_by(.data[[input$ar_group]]) %>%
        summarise(Cases = sum(.data[[input$ar_case]], na.rm = TRUE),
                  Population = sum(.data[[input$ar_pop]], na.rm = TRUE),
                  Attack_Rate = Cases / Population * 100,
                  Lower_CI = prop.test(Cases, Population, conf.level = input$ar_conf)$conf.int[1] * 100,
                  Upper_CI = prop.test(Cases, Population, conf.level = input$ar_conf)$conf.int[2] * 100)
    }
    
    datatable(ar_data, options = list(pageLength = 10))
  })
  
  output$download_ar <- downloadHandler(
    filename = function() {
      paste("attack_rate_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      df <- data()
      
      if(input$ar_group == "None") {
        ar_data <- df %>%
          summarise(Cases = sum(.data[[input$ar_case]], na.rm = TRUE),
                    Population = sum(.data[[input$ar_pop]], na.rm = TRUE),
                    Attack_Rate = Cases / Population * 100,
                    Lower_CI = prop.test(Cases, Population, conf.level = input$ar_conf)$conf.int[1] * 100,
                    Upper_CI = prop.test(Cases, Population, conf.level = input$ar_conf)$conf.int[2] * 100)
      } else {
        ar_data <- df %>%
          group_by(.data[[input$ar_group]]) %>%
          summarise(Cases = sum(.data[[input$ar_case]], na.rm = TRUE),
                    Population = sum(.data[[input$ar_pop]], na.rm = TRUE),
                    Attack_Rate = Cases / Population * 100,
                    Lower_CI = prop.test(Cases, Population, conf.level = input$ar_conf)$conf.int[1] * 100,
                    Upper_CI = prop.test(Cases, Population, conf.level = input$ar_conf)$conf.int[2] * 100)
      }
      
      ft <- flextable(ar_data) %>%
        set_caption("Attack Rate Analysis") %>%
        theme_zebra()
      
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Mortality rates
  output$mort_table <- renderDT({
    req(input$mort_case, input$mort_pop)
    df <- data()
    
    if(input$mort_measure == "Mortality Rate") {
      mort_data <- df %>%
        summarise(Deaths = sum(.data[[input$mort_case]], na.rm = TRUE),
                  Population = sum(.data[[input$mort_pop]], na.rm = TRUE),
                  Mortality_Rate = Deaths / Population * 1000)
    } else {
      mort_data <- df %>%
        summarise(Deaths = sum(.data[[input$mort_case]], na.rm = TRUE),
                  Cases = sum(.data[[input$mort_pop]], na.rm = TRUE),
                  CFR = Deaths / Cases * 100)
    }
    
    datatable(mort_data)
  })
  
  output$download_mort <- downloadHandler(
    filename = "mortality_table.docx",
    content = function(file) {
      df <- data()
      
      if(input$mort_measure == "Mortality Rate") {
        mort_data <- df %>%
          summarise(Deaths = sum(.data[[input$mort_case]], na.rm = TRUE),
                    Population = sum(.data[[input$mort_pop]], na.rm = TRUE),
                    Mortality_Rate = Deaths / Population * 1000)
      } else {
        mort_data <- df %>%
          summarise(Deaths = sum(.data[[input$mort_case]], na.rm = TRUE),
                    Cases = sum(.data[[input$mort_pop]], na.rm = TRUE),
                    CFR = Deaths / Cases * 100)
      }
      
      ft <- flextable(mort_data)
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Life table analysis
  output$lifetable_results <- renderDT({
    req(input$lt_time, input$lt_event)
    df <- data()
    
    # Parse time breaks
    breaks <- as.numeric(unlist(strsplit(input$lt_breaks, ",")))
    
    # Create survival object
    time <- as.numeric(df[[input$lt_time]])
    event <- as.numeric(df[[input$lt_event]])
    surv_obj <- Surv(time = time, event = event)
    
    # Create life table
    lt <- survfit(surv_obj ~ 1)
    lt_summary <- summary(lt, times = breaks, extend = TRUE)
    
    # Create life table dataframe
    lifetable <- data.frame(
      Time = lt_summary$time,
      AtRisk = lt_summary$n.risk,
      Events = lt_summary$n.event,
      Survival = lt_summary$surv,
      Lower_CI = lt_summary$lower,
      Upper_CI = lt_summary$upper
    )
    
    datatable(lifetable, options = list(scrollX = TRUE, pageLength = 10))
  })
  
  output$download_lifetable <- downloadHandler(
    filename = function() {
      paste("life_table_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      df <- data()
      breaks <- as.numeric(unlist(strsplit(input$lt_breaks, ",")))
      time <- as.numeric(df[[input$lt_time]])
      event <- as.numeric(df[[input$lt_event]])
      surv_obj <- Surv(time = time, event = event)
      lt <- survfit(surv_obj ~ 1)
      lt_summary <- summary(lt, times = breaks, extend = TRUE)
      
      lifetable <- data.frame(
        Time = lt_summary$time,
        AtRisk = lt_summary$n.risk,
        Events = lt_summary$n.event,
        Survival = lt_summary$surv,
        Lower_CI = lt_summary$lower,
        Upper_CI = lt_summary$upper
      )
      
      ft <- flextable(lifetable) %>%
        set_caption(input$lt_title) %>%
        theme_zebra()
      
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Risk ratio
  output$rr_results <- renderPrint({
    req(input$rr_outcome, input$rr_exposure)
    df <- data()
    
    # Create 2x2 table
    tab <- table(df[[input$rr_exposure]], df[[input$rr_outcome]])
    
    # Calculate measures
    a <- tab[2,2]  # Exposed cases
    b <- tab[2,1]  # Exposed non-cases
    c <- tab[1,2]  # Unexposed cases
    d <- tab[1,1]  # Unexposed non-cases
    
    risk_exposed <- a / (a + b)
    risk_unexposed <- c / (c + d)
    rr <- risk_exposed / risk_unexposed
    
    # Confidence interval
    se_log_rr <- sqrt(1/a - 1/(a+b) + 1/c - 1/(c+d))
    ci_lower <- exp(log(rr) - qnorm(1 - (1 - input$rr_conf)/2) * se_log_rr)
    ci_upper <- exp(log(rr) + qnorm(1 - (1 - input$rr_conf)/2) * se_log_rr)
    
    cat("2x2 Table:\n")
    print(tab)
    cat("\nRisk in", input$rr_exposed_label, "group:", round(risk_exposed, 4), "\n")
    cat("Risk in", input$rr_unexposed_label, "group:", round(risk_unexposed, 4), "\n")
    cat("Risk Ratio:", round(rr, 4), "\n")
    cat(input$rr_conf*100, "% Confidence Interval:", round(ci_lower, 4), "-", round(ci_upper, 4), "\n")
  })
  
  output$rr_table <- renderDT({
    req(input$rr_outcome, input$rr_exposure)
    df <- data()
    
    # Create 2x2 table
    tab <- table(df[[input$rr_exposure]], df[[input$rr_outcome]])
    
    # Calculate measures
    a <- tab[2,2]  # Exposed cases
    b <- tab[2,1]  # Exposed non-cases
    c <- tab[1,2]  # Unexposed cases
    d <- tab[1,1]  # Unexposed non-cases
    
    risk_exposed <- a / (a + b)
    risk_unexposed <- c / (c + d)
    rr <- risk_exposed / risk_unexposed
    
    # Confidence interval
    se_log_rr <- sqrt(1/a - 1/(a+b) + 1/c - 1/(c+d))
    ci_lower <- exp(log(rr) - qnorm(1 - (1 - input$rr_conf)/2) * se_log_rr)
    ci_upper <- exp(log(rr) + qnorm(1 - (1 - input$rr_conf)/2) * se_log_rr)
    
    rr_data <- data.frame(
      Measure = c(paste("Risk in", input$rr_exposed_label, "group"), 
                  paste("Risk in", input$rr_unexposed_label, "group"), 
                  "Risk Ratio", 
                  paste(input$rr_conf*100, "% CI Lower"), 
                  paste(input$rr_conf*100, "% CI Upper")),
      Value = c(risk_exposed, risk_unexposed, rr, ci_lower, ci_upper)
    )
    
    datatable(rr_data)
  })
  
  output$download_rr <- downloadHandler(
    filename = "risk_ratio_table.docx",
    content = function(file) {
      df <- data()
      
      tab <- table(df[[input$rr_exposure]], df[[input$rr_outcome]])
      
      a <- tab[2,2]
      b <- tab[2,1]
      c <- tab[1,2]
      d <- tab[1,1]
      
      risk_exposed <- a / (a + b)
      risk_unexposed <- c / (c + d)
      rr <- risk_exposed / risk_unexposed
      
      se_log_rr <- sqrt(1/a - 1/(a+b) + 1/c - 1/(c+d))
      ci_lower <- exp(log(rr) - qnorm(1 - (1 - input$rr_conf)/2) * se_log_rr)
      ci_upper <- exp(log(rr) + qnorm(1 - (1 - input$rr_conf)/2) * se_log_rr)
      
      rr_data <- data.frame(
        Measure = c(paste("Risk in", input$rr_exposed_label, "group"), 
                    paste("Risk in", input$rr_unexposed_label, "group"), 
                    "Risk Ratio", 
                    paste(input$rr_conf*100, "% CI Lower"), 
                    paste(input$rr_conf*100, "% CI Upper")),
        Value = c(risk_exposed, risk_unexposed, rr, ci_lower, ci_upper)
      )
      
      ft <- flextable(rr_data)
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Odds ratio
  output$or_results <- renderPrint({
    req(input$or_outcome, input$or_exposure)
    df <- data()
    
    # Create 2x2 table
    tab <- table(df[[input$or_exposure]], df[[input$or_outcome]])
    
    # Calculate measures
    a <- tab[2,2]  # Exposed cases
    b <- tab[2,1]  # Exposed non-cases
    c <- tab[1,2]  # Unexposed cases
    d <- tab[1,1]  # Unexposed non-cases
    
    or <- (a * d) / (b * c)
    
    # Confidence interval
    se_log_or <- sqrt(1/a + 1/b + 1/c + 1/d)
    ci_lower <- exp(log(or) - qnorm(1 - (1 - input$or_conf)/2) * se_log_or)
    ci_upper <- exp(log(or) + qnorm(1 - (1 - input$or_conf)/2) * se_log_or)
    
    cat("2x2 Table:\n")
    print(tab)
    cat("\nOdds Ratio:", round(or, 4), "\n")
    cat(input$or_conf*100, "% Confidence Interval:", round(ci_lower, 4), "-", round(ci_upper, 4), "\n")
  })
  
  output$or_table <- renderDT({
    req(input$or_outcome, input$or_exposure)
    df <- data()
    
    # Create 2x2 table
    tab <- table(df[[input$or_exposure]], df[[input$or_outcome]])
    
    # Calculate measures
    a <- tab[2,2]  # Exposed cases
    b <- tab[2,1]  # Exposed non-cases
    c <- tab[1,2]  # Unexposed cases
    d <- tab[1,1]  # Unexposed non-cases
    
    or <- (a * d) / (b * c)
    
    # Confidence interval
    se_log_or <- sqrt(1/a + 1/b + 1/c + 1/d)
    ci_lower <- exp(log(or) - qnorm(1 - (1 - input$or_conf)/2) * se_log_or)
    ci_upper <- exp(log(or) + qnorm(1 - (1 - input$or_conf)/2) * se_log_or)
    
    or_data <- data.frame(
      Measure = c("Odds Ratio", 
                  paste(input$or_conf*100, "% CI Lower"), 
                  paste(input$or_conf*100, "% CI Upper")),
      Value = c(or, ci_lower, ci_upper)
    )
    
    datatable(or_data)
  })
  
  output$download_or <- downloadHandler(
    filename = "odds_ratio_table.docx",
    content = function(file) {
      df <- data()
      
      tab <- table(df[[input$or_exposure]], df[[input$or_outcome]])
      
      a <- tab[2,2]
      b <- tab[2,1]
      c <- tab[1,2]
      d <- tab[1,1]
      
      or <- (a * d) / (b * c)
      
      se_log_or <- sqrt(1/a + 1/b + 1/c + 1/d)
      ci_lower <- exp(log(or) - qnorm(1 - (1 - input$or_conf)/2) * se_log_or)
      ci_upper <- exp(log(or) + qnorm(1 - (1 - input$or_conf)/2) * se_log_or)
      
      or_data <- data.frame(
        Measure = c("Odds Ratio", 
                    paste(input$or_conf*100, "% CI Lower"), 
                    paste(input$or_conf*100, "% CI Upper")),
        Value = c(or, ci_lower, ci_upper)
      )
      
      ft <- flextable(or_data)
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Attributable risk
  output$ar_results <- renderPrint({
    req(input$ar_outcome, input$ar_exposure)
    df <- data()
    
    # Create 2x2 table
    tab <- table(df[[input$ar_exposure]], df[[input$ar_outcome]])
    
    # Calculate measures
    a <- tab[2,2]  # Exposed cases
    b <- tab[2,1]  # Exposed non-cases
    c <- tab[1,2]  # Unexposed cases
    d <- tab[1,1]  # Unexposed non-cases
    
    risk_exposed <- a / (a + b)
    risk_unexposed <- c / (c + d)
    ar <- risk_exposed - risk_unexposed
    
    if(input$ar_paf) {
      paf <- (ar / risk_exposed) * 100
    }
    
    # Confidence interval for AR
    var_ar <- (risk_exposed * (1 - risk_exposed)) / (a + b) + 
      (risk_unexposed * (1 - risk_unexposed)) / (c + d)
    se_ar <- sqrt(var_ar)
    ci_lower <- ar - qnorm(1 - (1 - input$ar_conf)/2) * se_ar
    ci_upper <- ar + qnorm(1 - (1 - input$ar_conf)/2) * se_ar
    
    cat("2x2 Table:\n")
    print(tab)
    cat("\nRisk in", input$ar_exposed_label, "group:", round(risk_exposed, 4), "\n")
    cat("Risk in", input$ar_unexposed_label, "group:", round(risk_unexposed, 4), "\n")
    cat("Attributable Risk:", round(ar, 4), "\n")
    if(input$ar_paf) {
      cat("Population Attributable Fraction:", round(paf, 2), "%\n")
    }
    cat(input$ar_conf*100, "% Confidence Interval for AR:", round(ci_lower, 4), "-", round(ci_upper, 4), "\n")
  })
  
  output$ar_table <- renderDT({
    req(input$ar_outcome, input$ar_exposure)
    df <- data()
    
    # Create 2x2 table
    tab <- table(df[[input$ar_exposure]], df[[input$ar_outcome]])
    
    # Calculate measures
    a <- tab[2,2]  # Exposed cases
    b <- tab[2,1]  # Exposed non-cases
    c <- tab[1,2]  # Unexposed cases
    d <- tab[1,1]  # Unexposed non-cases
    
    risk_exposed <- a / (a + b)
    risk_unexposed <- c / (c + d)
    ar <- risk_exposed - risk_unexposed
    
    if(input$ar_paf) {
      paf <- (ar / risk_exposed) * 100
    }
    
    # Confidence interval for AR
    var_ar <- (risk_exposed * (1 - risk_exposed)) / (a + b) + 
      (risk_unexposed * (1 - risk_unexposed)) / (c + d)
    se_ar <- sqrt(var_ar)
    ci_lower <- ar - qnorm(1 - (1 - input$ar_conf)/2) * se_ar
    ci_upper <- ar + qnorm(1 - (1 - input$ar_conf)/2) * se_ar
    
    if(input$ar_paf) {
      ar_data <- data.frame(
        Measure = c(paste("Risk in", input$ar_exposed_label, "group"), 
                    paste("Risk in", input$ar_unexposed_label, "group"), 
                    "Attributable Risk", 
                    "Attributable Risk % (PAF)", 
                    paste(input$ar_conf*100, "% CI Lower (AR)"), 
                    paste(input$ar_conf*100, "% CI Upper (AR)")),
        Value = c(risk_exposed, risk_unexposed, ar, paf, ci_lower, ci_upper)
      )
    } else {
      ar_data <- data.frame(
        Measure = c(paste("Risk in", input$ar_exposed_label, "group"), 
                    paste("Risk in", input$ar_unexposed_label, "group"), 
                    "Attributable Risk", 
                    paste(input$ar_conf*100, "% CI Lower"), 
                    paste(input$ar_conf*100, "% CI Upper")),
        Value = c(risk_exposed, risk_unexposed, ar, ci_lower, ci_upper)
      )
    }
    
    datatable(ar_data)
  })
  
  output$download_arisk <- downloadHandler(
    filename = "attributable_risk_table.docx",
    content = function(file) {
      df <- data()
      
      tab <- table(df[[input$ar_exposure]], df[[input$ar_outcome]])
      
      a <- tab[2,2]
      b <- tab[2,1]
      c <- tab[1,2]
      d <- tab[1,1]
      
      risk_exposed <- a / (a + b)
      risk_unexposed <- c / (c + d)
      ar <- risk_exposed - risk_unexposed
      
      if(input$ar_paf) {
        paf <- (ar / risk_exposed) * 100
      }
      
      var_ar <- (risk_exposed * (1 - risk_exposed)) / (a + b) + 
        (risk_unexposed * (1 - risk_unexposed)) / (c + d)
      se_ar <- sqrt(var_ar)
      ci_lower <- ar - qnorm(1 - (1 - input$ar_conf)/2) * se_ar
      ci_upper <- ar + qnorm(1 - (1 - input$ar_conf)/2) * se_ar
      
      if(input$ar_paf) {
        ar_data <- data.frame(
          Measure = c(paste("Risk in", input$ar_exposed_label, "group"), 
                      paste("Risk in", input$ar_unexposed_label, "group"), 
                      "Attributable Risk", 
                      "Attributable Risk % (PAF)", 
                      paste(input$ar_conf*100, "% CI Lower (AR)"), 
                      paste(input$ar_conf*100, "% CI Upper (AR)")),
          Value = c(risk_exposed, risk_unexposed, ar, paf, ci_lower, ci_upper)
        )
      } else {
        ar_data <- data.frame(
          Measure = c(paste("Risk in", input$ar_exposed_label, "group"), 
                      paste("Risk in", input$ar_unexposed_label, "group"), 
                      "Attributable Risk", 
                      paste(input$ar_conf*100, "% CI Lower"), 
                      paste(input$ar_conf*100, "% CI Upper")),
          Value = c(risk_exposed, risk_unexposed, ar, ci_lower, ci_upper)
        )
      }
      
      ft <- flextable(ar_data)
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Poisson regression with reference category selection using glm()
  output$poisson_ref_ui <- renderUI({
    req(input$poisson_vars)
    df <- data()
    
    ref_ui <- tagList()
    
    for (var in input$poisson_vars) {
      if (is.factor(df[[var]]) || is.character(df[[var]])) {
        levels <- if (is.factor(df[[var]])) levels(df[[var]]) else unique(df[[var]])
        ref_ui <- tagAppendChild(
          ref_ui,
          selectInput(paste0("poisson_ref_", var), 
                      paste("Reference category for", var), 
                      choices = levels)
        )
      }
    }
    
    ref_ui
  })
  
  output$poisson_summary <- renderPrint({
    req(input$poisson_outcome, input$poisson_offset, input$poisson_vars)
    df <- data()
    
    # Prepare formula
    formula <- paste(input$poisson_outcome, "~", paste(input$poisson_vars, collapse = " + "))
    
    # Set reference categories for factors
    for (var in input$poisson_vars) {
      if (is.factor(df[[var]])) {
        ref <- input[[paste0("poisson_ref_", var)]]
        if (!is.null(ref)) {
          df[[var]] <- relevel(df[[var]], ref = ref)
        }
      } else if (is.character(df[[var]])) {
        ref <- input[[paste0("poisson_ref_", var)]]
        if (!is.null(ref)) {
          df[[var]] <- factor(df[[var]])
          df[[var]] <- relevel(df[[var]], ref = ref)
        }
      }
    }
    
    # Fit Poisson model using glm()
    poisson_fit <- glm(as.formula(formula), 
                       family = poisson(link = "log"), 
                       data = df,
                       offset = log(df[[input$poisson_offset]]))
    
    cat("Poisson Regression Results:\n\n")
    print(summary(poisson_fit))
    
    if (input$poisson_irr) {
      cat("\nIncidence Rate Ratios (IRR):\n")
      print(exp(cbind(IRR = coef(poisson_fit), confint(poisson_fit, level = input$poisson_conf))))
    }
  })
  
  output$poisson_table <- renderDT({
    req(input$poisson_outcome, input$poisson_offset, input$poisson_vars)
    df <- data()
    
    # Prepare formula
    formula <- paste(input$poisson_outcome, "~", paste(input$poisson_vars, collapse = " + "))
    
    # Set reference categories for factors
    for (var in input$poisson_vars) {
      if (is.factor(df[[var]])) {
        ref <- input[[paste0("poisson_ref_", var)]]
        if (!is.null(ref)) {
          df[[var]] <- relevel(df[[var]], ref = ref)
        }
      } else if (is.character(df[[var]])) {
        ref <- input[[paste0("poisson_ref_", var)]]
        if (!is.null(ref)) {
          df[[var]] <- factor(df[[var]])
          df[[var]] <- relevel(df[[var]], ref = ref)
        }
      }
    }
    
    # Fit Poisson model using glm()
    poisson_fit <- glm(as.formula(formula), 
                       family = poisson(link = "log"), 
                       data = df,
                       offset = log(df[[input$poisson_offset]]))
    
    # Create results table using tidy approach
    if (input$poisson_irr) {
      results <- broom::tidy(poisson_fit, conf.int = TRUE, conf.level = input$poisson_conf, exponentiate = TRUE) %>%
        select(term, estimate, conf.low, conf.high) %>%
        rename(IRR = estimate, `Lower CI` = conf.low, `Upper CI` = conf.high)
    } else {
      results <- broom::tidy(poisson_fit, conf.int = TRUE, conf.level = input$poisson_conf) %>%
        select(term, estimate, conf.low, conf.high) %>%
        rename(Coefficient = estimate, `Lower CI` = conf.low, `Upper CI` = conf.high)
    }
    
    datatable(results, options = list(pageLength = 10))
  })
  
  output$download_poisson <- downloadHandler(
    filename = function() {
      paste("poisson_regression_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      df <- data()
      
      # Prepare formula
      formula <- paste(input$poisson_outcome, "~", paste(input$poisson_vars, collapse = " + "))
      
      # Set reference categories for factors
      for (var in input$poisson_vars) {
        if (is.factor(df[[var]]) || is.character(df[[var]])) {
          ref <- input[[paste0("poisson_ref_", var)]]
          if (!is.null(ref)) {
            df[[var]] <- factor(df[[var]])
            df[[var]] <- relevel(df[[var]], ref = ref)
          }
        }
      }
      
      # Fit Poisson model using glm()
      poisson_fit <- glm(as.formula(formula), 
                         family = poisson(link = "log"), 
                         data = df,
                         offset = log(df[[input$poisson_offset]]))
      
      # Create results table
      results <- tidy(poisson_fit, conf.int = TRUE, conf.level = input$poisson_conf) %>%
        mutate(
          estimate = if (input$poisson_irr) exp(estimate) else estimate,
          conf.low = if (input$poisson_irr) exp(conf.low) else conf.low,
          conf.high = if (input$poisson_irr) exp(conf.high) else conf.high
        ) %>%
        select(term, estimate, conf.low, conf.high) %>%
        rename(
          !!if (input$poisson_irr) "IRR" else "Coefficient" := estimate,
          `Lower CI` = conf.low,
          `Upper CI` = conf.high
        )
      
      ft <- flextable(results) %>%
        set_caption("Poisson Regression Results") %>%
        theme_zebra()
      
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Create survival object function
  create_surv_obj <- function(time_var, event_var, df) {
    time <- df[[time_var]]
    event <- df[[event_var]]
    
    # Convert event to numeric if needed
    if (is.character(event) || is.factor(event)) {
      event <- as.numeric(as.factor(event)) - 1
    } else if (is.logical(event)) {
      event <- as.numeric(event)
    }
    
    # Ensure event is binary (0/1)
    event <- ifelse(event > 0, 1, 0)
    
    Surv(time = time, event = event)
  }
  
  # Kaplan-Meier analysis
  output$km_plot <- renderPlot({
    req(input$km_time, input$km_event)
    df <- data()
    
    # Check if manual entry is provided
    if (input$km_time_entry != "" && input$km_event_entry != "") {
      # Use manual entry data
      time <- as.numeric(unlist(strsplit(input$km_time_entry, ",")))
      event <- as.numeric(unlist(strsplit(input$km_event_entry, ",")))
      df <- data.frame(time = time, event = event)
      surv_obj <- Surv(time = df$time, event = df$event)
    } else {
      # Use dataset variables with robust survival object creation
      surv_obj <- tryCatch({
        create_surv_obj(input$km_time, input$km_event, df)
      }, error = function(e) {
        # Fallback to simple creation if specialized function fails
        time <- df[[input$km_time]]
        event <- df[[input$km_event]]
        Surv(time = time, event = as.numeric(event))
      })
    }
    
    if(input$km_group == "None") {
      fit <- survfit(surv_obj ~ 1, data = df)
      p <- ggsurvplot(fit, data = df, 
                      conf.int = input$km_conf, 
                      risk.table = input$km_risktable,
                      title = input$km_title,
                      font.title = input$km_title_size,
                      font.x = input$km_x_size,
                      font.y = input$km_y_size,
                      ggtheme = theme_minimal())
    } else {
      fit <- survfit(surv_obj ~ df[[input$km_group]], data = df)
      p <- ggsurvplot(fit, data = df, 
                      conf.int = input$km_conf, 
                      risk.table = input$km_risktable,
                      legend.title = input$km_group,
                      title = input$km_title,
                      font.title = input$km_title_size,
                      font.x = input$km_x_size,
                      font.y = input$km_y_size,
                      palette = c(input$km_color1, input$km_color2),
                      ggtheme = theme_minimal())
    }
    
    print(p)
  })
  
  output$download_km <- downloadHandler(
    filename = "kaplan_meier_plot.png",
    content = function(file) {
      df <- data()
      
      # Check if manual entry is provided
      if (input$km_time_entry != "" && input$km_event_entry != "") {
        # Use manual entry data
        time <- as.numeric(unlist(strsplit(input$km_time_entry, ",")))
        event <- as.numeric(unlist(strsplit(input$km_event_entry, ",")))
        df <- data.frame(time = time, event = event)
        surv_obj <- Surv(time = df$time, event = df$event)
      } else {
        # Use dataset variables with robust survival object creation
        surv_obj <- tryCatch({
          create_surv_obj(input$km_time, input$km_event, df)
        }, error = function(e) {
          # Fallback to simple creation if specialized function fails
          time <- df[[input$km_time]]
          event <- df[[input$km_event]]
          Surv(time = time, event = as.numeric(event))
        })
      }
      
      if(input$km_group == "None") {
        fit <- survfit(surv_obj ~ 1, data = df)
        p <- ggsurvplot(fit, data = df, 
                        conf.int = input$km_conf, 
                        risk.table = input$km_risktable,
                        title = input$km_title,
                        font.title = input$km_title_size,
                        font.x = input$km_x_size,
                        font.y = input$km_y_size,
                        ggtheme = theme_minimal())
      } else {
        fit <- survfit(surv_obj ~ df[[input$km_group]], data = df)
        p <- ggsurvplot(fit, data = df, 
                        conf.int = input$km_conf, 
                        risk.table = input$km_risktable,
                        legend.title = input$km_group,
                        title = input$km_title,
                        font.title = input$km_title_size,
                        font.x = input$km_x_size,
                        font.y = input$km_y_size,
                        palette = c(input$km_color1, input$km_color2),
                        ggtheme = theme_minimal())
      }
      
      ggsave(file, plot = print(p), device = "png", width = 10, height = 6)
    }
  )
  
  # Log-rank test
  output$logrank_results <- renderPrint({
    req(input$lr_time, input$lr_event, input$lr_group)
    df <- data()
    
    surv_obj <- tryCatch({
      create_surv_obj(input$lr_time, input$lr_event, df)
    }, error = function(e) {
      # Fallback to simple creation if specialized function fails
      time <- df[[input$lr_time]]
      event <- df[[input$lr_event]]
      Surv(time = time, event = as.numeric(event))
    })
    
    fit <- survdiff(surv_obj ~ df[[input$lr_group]])
    
    cat("Log-Rank Test Results:\n")
    print(fit)
  })
  
  output$download_logrank <- downloadHandler(
    filename = "logrank_test.docx",
    content = function(file) {
      df <- data()
      
      surv_obj <- tryCatch({
        create_surv_obj(input$lr_time, input$lr_event, df)
      }, error = function(e) {
        # Fallback to simple creation if specialized function fails
        time <- df[[input$lr_time]]
        event <- df[[input$lr_event]]
        Surv(time = time, event = as.numeric(event))
      })
      
      fit <- survdiff(surv_obj ~ df[[input$lr_group]])
      
      # Convert to flextable
      res <- capture.output(fit)
      ft <- flextable(data.frame(Results = res))
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Cox regression
  output$cox_summary <- renderPrint({
    req(input$cox_time, input$cox_event, input$cox_vars)
    df <- data()
    
    # Create survival object
    surv_obj <- tryCatch({
      create_surv_obj(input$cox_time, input$cox_event, df)
    }, error = function(e) {
      # Fallback to simple creation if specialized function fails
      time <- df[[input$cox_time]]
      event <- df[[input$cox_event]]
      Surv(time = time, event = as.numeric(event))
    })
    
    # Create formula
    formula <- as.formula(paste("surv_obj ~", paste(input$cox_vars, collapse = " + ")))
    
    # Fit Cox model
    cox_fit <- coxph(formula, data = df)
    summary(cox_fit, conf.int = input$cox_conf)
  })
  
  output$download_cox <- downloadHandler(
    filename = "cox_regression.docx",
    content = function(file) {
      df <- data()
      
      # Create survival object
      surv_obj <- tryCatch({
        create_surv_obj(input$cox_time, input$cox_event, df)
      }, error = function(e) {
        # Fallback to simple creation if specialized function fails
        time <- df[[input$cox_time]]
        event <- df[[input$cox_event]]
        Surv(time = time, event = as.numeric(event))
      })
      
      formula <- as.formula(paste("surv_obj ~", paste(input$cox_vars, collapse = " + ")))
      cox_fit <- coxph(formula, data = df)
      
      # Convert summary to flextable
      sum_cox <- summary(cox_fit, conf.int = input$cox_conf)
      cox_table <- as.data.frame(sum_cox$coefficients)
      cox_table <- cbind(Variable = rownames(cox_table), cox_table)
      rownames(cox_table) <- NULL
      
      ft <- flextable(cox_table) %>%
        set_caption("Cox Regression Results") %>%
        theme_zebra()
      
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
}

# Run the application
shinyApp(ui = ui, server = server)
