library(shiny)
library(shinydashboard)
library(ggplot2)
library(dplyr)
library(ggpubr)
library(survival)
library(survminer)
library(officer)
library(DT)
library(haven)
library(readxl)
library(plotly)
library(colourpicker)
library(epitools)
library(tidyverse)
library(broom)
library(car)
library(emmeans)
library(fresh)
library(lm.beta)
library(MASS)
library(nnet)
library(dunn.test)

# Helper function for survival object creation
create_surv_obj <- function(time_var, event_var, data) {
  time <- data[[time_var]]
  event <- data[[event_var]]
  
  # Convert event to numeric if it's not
  if (!is.numeric(event)) {
    event <- as.numeric(as.factor(event)) - 1
  }
  
  Surv(time = time, event = event)
}

# UI
ui <- navbarPage(
  title = div(
    img(src = "", 
        height = "30", 
        style = "margin-right:10px;padding-bottom:3px;"),
    "📈 EpiDem Suite™"
  ),
  windowTitle = "EpiDem Suite - Epidemiological Analysis Tool",
  id = "nav",
  theme = bslib::bs_theme(version = 5, 
                          bootswatch = "flatly",
                          primary = "#0C5EA8",   # CDC Blue
                          secondary = "#CD2026", # CDC Red
                          success = "#5CB85C",
                          font_scale = 0.95),
  
  # Home/About Tab
  tabPanel("Home", icon = icon("house"),
           div(class = "jumbotron",
               style = "background: linear-gradient(rgba(255,255,255,0.9), rgba(255,255,255,0.9)), 
                        url('');
                        background-size: cover; 
                        padding: 3rem 2rem; 
                        margin-bottom: 2rem; 
                        border-radius: 0.3rem;",
               h1("EpiDem Suite", class = "display-4"),
               p("A comprehensive epidemiological analysis platform for public health professionals.", 
                 class = "lead"),
               hr(class = "my-4"),
               p("Upload your dataset and utilize the tools in the navigation bar to perform descriptive analyses, 
                 calculate disease frequency measures, assess statistical associations, and build regression models."),
               p("This tool follows established guidelines for epidemiological analysis. Developed by Mudasir Mohammed Ibrahim (BSc, RN)"),
               actionButton("goto_data", "Get Started →", 
                            class = "btn-primary btn-lg", 
                            icon = icon("rocket"))
           ),
           
           fluidRow(
             column(4,
                    div(class = "card",
                        div(class = "card-body",
                            h3("Descriptive Analysis", class = "card-title"),
                            p("Explore frequency distributions, summary statistics, and create epidemic curves to understand your data's basic characteristics."),
                            actionButton("goto_descriptive", "Learn More", class = "btn-outline-primary")
                        )
                    )
             ),
             column(4,
                    div(class = "card",
                        div(class = "card-body",
                            h3("Disease Metrics", class = "card-title"),
                            p("Calculate incidence rates, prevalence, attack rates, and mortality measures to quantify disease burden in populations."),
                            actionButton("goto_disease", "Learn More", class = "btn-outline-primary")
                        )
                    )
             ),
             column(4,
                    div(class = "card",
                        div(class = "card-body",
                            h3("Statistical Tools", class = "card-title"),
                            p("Perform association tests, regression analyses, and survival analysis to identify risk factors and measure effects."),
                            actionButton("goto_statistics", "Learn More", class = "btn-outline-primary")
                        )
                    )
             )
           )
  ),
  
  # Data Upload Tab
  tabPanel("Data Upload", icon = icon("database"),
           sidebarLayout(
             sidebarPanel(
               fileInput("file1", "Choose File",
                         accept = c(".csv", ".xls", ".xlsx", ".sav", ".dta")),
               tags$hr(),
               checkboxInput("header", "Header", TRUE),
               selectInput("sep", "Separator",
                           choices = c(Comma = ",", Semicolon = ";", Tab = "\t"),
                           selected = ","),
               selectInput("quote", "Quote",
                           choices = c(None = "", "Double Quote" = '"', "Single Quote" = "'"),
                           selected = '"'),
               conditionalPanel(
                 condition = "input.file1 != null && (input.file1.type == 'xls' || input.file1.type == 'xlsx')",
                 numericInput("sheet", "Sheet Number", value = 1)
               )
             ),
             mainPanel(
               DTOutput("contents")
             )
           )
  ),
  
  # Descriptive Analysis Tab
  tabPanel("Descriptive Analysis", icon = icon("chart-bar"),
           tabsetPanel(
             tabPanel("Frequency Analysis",
                      sidebarLayout(
                        sidebarPanel(
                          selectInput("freq_var", "Select Variable:", choices = NULL),
                          colourInput("freq_color", "Bar Color:", value = "#0C5EA8"),
                          textInput("freq_title", "Plot Title:", value = "Frequency Distribution"),
                          numericInput("freq_title_size", "Title Size:", value = 16),
                          numericInput("freq_x_size", "X-axis Label Size:", value = 14),
                          numericInput("freq_y_size", "Y-axis Label Size:", value = 14),
                          radioButtons("freq_type", "Download Format:",
                                       choices = c("Table", "Plot"), selected = "Table"),
                          downloadButton("download_freq", "Download Results")
                        ),
                        mainPanel(
                          DTOutput("freq_table"),
                          plotlyOutput("freq_plot")
                        )
                      )
             ),
             tabPanel("Central Tendency",
                      sidebarLayout(
                        sidebarPanel(
                          selectInput("central_var", "Select Numeric Variable:", choices = NULL),
                          colourInput("central_color", "Plot Color:", value = "#0C5EA8"),
                          textInput("central_title", "Plot Title:", value = "Central Tendency"),
                          numericInput("central_bins", "Number of Bins:", value = 30),
                          numericInput("central_title_size", "Title Size:", value = 16),
                          numericInput("central_x_size", "X-axis Label Size:", value = 14),
                          numericInput("central_y_size", "Y-axis Label Size:", value = 14),
                          radioButtons("central_type", "Download Format:",
                                       choices = c("Table", "Boxplot", "Histogram"), selected = "Table"),
                          downloadButton("download_central", "Download Results")
                        ),
                        mainPanel(
                          DTOutput("central_table"),
                          plotlyOutput("central_boxplot"),
                          plotlyOutput("central_histogram")
                        )
                      )
             ),
             tabPanel("Epidemic Curve",
                      sidebarLayout(
                        sidebarPanel(
                          selectInput("epi_date", "Date Variable:", choices = NULL),
                          selectInput("epi_case", "Case Count Variable:", choices = NULL),
                          selectInput("epi_group", "Grouping Variable (Optional):", choices = NULL),
                          selectInput("epi_interval", "Time Interval:",
                                      choices = c("Day", "Week", "Month", "Year"), selected = "Week"),
                          colourInput("epi_color", "Bar Color:", value = "#CD2026"),
                          textInput("epi_title", "Plot Title:", value = "Epidemic Curve"),
                          numericInput("epi_title_size", "Title Size:", value = 16),
                          numericInput("epi_x_size", "X-axis Label Size:", value = 14),
                          numericInput("epi_y_size", "Y-axis Label Size:", value = 14),
                          downloadButton("download_epicurve", "Download Plot")
                        ),
                        mainPanel(
                          plotlyOutput("epi_curve")
                        )
                      )
             ),
             tabPanel("Spatial Analysis",
                      sidebarLayout(
                        sidebarPanel(
                          selectInput("map_var", "Variable to Map:", choices = NULL),
                          selectInput("map_region", "Region Variable:", choices = NULL),
                          selectInput("map_type", "Map Type:",
                                      choices = c("Choropleth", "Point Map"), selected = "Choropleth"),
                          colourInput("map_color", "Point Color:", value = "#0C5EA8"),
                          textInput("map_title", "Plot Title:", value = "Spatial Distribution"),
                          numericInput("map_title_size", "Title Size:", value = 16),
                          numericInput("map_x_size", "X-axis Label Size:", value = 14),
                          numericInput("map_y_size", "Y-axis Label Size:", value = 14),
                          downloadButton("download_map", "Download Plot")
                        ),
                        mainPanel(
                          plotlyOutput("map_plot")
                        )
                      )
             ),
             tabPanel("Demographic Breakdown",
                      sidebarLayout(
                        sidebarPanel(
                          selectInput("demo_var", "Analysis Variable:", choices = NULL),
                          selectizeInput("demo_group", "Grouping Variables:", choices = NULL, multiple = TRUE),
                          downloadButton("download_demo", "Download Table")
                        ),
                        mainPanel(
                          DTOutput("demo_table")
                        )
                      )
             )
           )
  ),
  
  # Disease Frequency Tab
  tabPanel("Disease Frequency", icon = icon("viruses"),
           tabsetPanel(
             tabPanel("Incidence/Prevalence",
                      sidebarLayout(
                        sidebarPanel(
                          selectInput("ip_measure", "Select Measure:",
                                      choices = c("Incidence Rate", "Cumulative Incidence", "Prevalence"),
                                      selected = "Incidence Rate"),
                          selectInput("ip_case", "Case Count Variable:", choices = NULL),
                          selectInput("ip_pop", "Population Variable:", choices = NULL),
                          conditionalPanel(
                            condition = "input.ip_measure == 'Incidence Rate'",
                            selectInput("ip_time", "Person-Time Variable:", choices = NULL)
                          ),
                          downloadButton("download_ip", "Download Results")
                        ),
                        mainPanel(
                          DTOutput("ip_table")
                        )
                      )
             ),
             tabPanel("Attack Rates",
                      sidebarLayout(
                        sidebarPanel(
                          selectInput("ar_case", "Case Count Variable:", choices = NULL),
                          selectInput("ar_pop", "Population Variable:", choices = NULL),
                          selectInput("ar_group", "Grouping Variable (Optional):", choices = NULL),
                          radioButtons("ar_format", "Output Format:",
                                       choices = c("Table", "Detailed with CI"), selected = "Table"),
                          downloadButton("download_ar", "Download Results")
                        ),
                        mainPanel(
                          DTOutput("ar_table"),
                          DTOutput("ar_detailed_table")
                        )
                      )
             ),
             tabPanel("Mortality Rates",
                      sidebarLayout(
                        sidebarPanel(
                          selectInput("mort_measure", "Select Measure:",
                                      choices = c("Mortality Rate", "Case Fatality Rate"),
                                      selected = "Mortality Rate"),
                          selectInput("mort_case", "Death Count Variable:", choices = NULL),
                          selectInput("mort_pop", "Population Variable:", choices = NULL),
                          downloadButton("download_mort", "Download Results")
                        ),
                        mainPanel(
                          DTOutput("mort_table")
                        )
                      )
             ),
             tabPanel("Life Table Analysis",
                      sidebarLayout(
                        sidebarPanel(
                          selectInput("lt_time", "Time Variable:", choices = NULL),
                          selectInput("lt_event", "Event Variable:", choices = NULL),
                          textInput("lt_breaks", "Time Points (comma-separated):", value = "0,1,2,3,4,5"),
                          textInput("lt_title", "Table Title:", value = "Life Table Analysis"),
                          downloadButton("download_lifetable", "Download Table")
                        ),
                        mainPanel(
                          DTOutput("lifetable_results")
                        )
                      )
             )
           )
  ),
  
  # Statistical Associations Tab
  tabPanel("Statistical Associations", icon = icon("calculator"),
           tabsetPanel(
             tabPanel("Risk Ratios",
                      sidebarLayout(
                        sidebarPanel(
                          selectInput("rr_outcome", "Outcome Variable:", choices = NULL),
                          selectInput("rr_exposure", "Exposure Variable:", choices = NULL),
                          sliderInput("rr_conf", "Confidence Level:", min = 0.90, max = 0.99, value = 0.95, step = 0.01),
                          downloadButton("download_rr", "Download Results")
                        ),
                        mainPanel(
                          verbatimTextOutput("rr_results"),
                          DTOutput("rr_table")
                        )
                      )
             ),
             tabPanel("Odds Ratios",
                      sidebarLayout(
                        sidebarPanel(
                          selectInput("or_outcome", "Outcome Variable:", choices = NULL),
                          selectInput("or_exposure", "Exposure Variable:", choices = NULL),
                          sliderInput("or_conf", "Confidence Level:", min = 0.90, max = 0.99, value = 0.95, step = 0.01),
                          downloadButton("download_or", "Download Results")
                        ),
                        mainPanel(
                          verbatimTextOutput("or_results"),
                          DTOutput("or_table")
                        )
                      )
             ),
             tabPanel("Attributable Risk",
                      sidebarLayout(
                        sidebarPanel(
                          selectInput("ar_outcome", "Outcome Variable:", choices = NULL),
                          selectInput("ar_exposure", "Exposure Variable:", choices = NULL),
                          sliderInput("ar_conf", "Confidence Level:", min = 0.90, max = 0.99, value = 0.95, step = 0.01),
                          checkboxInput("ar_paf", "Calculate Population Attributable Fraction", FALSE),
                          downloadButton("download_arisk", "Download Results")
                        ),
                        mainPanel(
                          verbatimTextOutput("ar_results"),
                          DTOutput("ar_table")
                        )
                      )
             ),
             tabPanel("Poisson Regression",
                      sidebarLayout(
                        sidebarPanel(
                          selectInput("poisson_outcome", "Outcome Variable (Count):", choices = NULL),
                          selectInput("poisson_offset", "Offset Variable (Optional):", choices = NULL),
                          selectizeInput("poisson_vars", "Predictor Variables:", choices = NULL, multiple = TRUE),
                          uiOutput("poisson_ref_ui"),
                          sliderInput("poisson_conf", "Confidence Level:", min = 0.90, max = 0.99, value = 0.95, step = 0.01),
                          checkboxInput("poisson_irr", "Exponentiate Coefficients (IRR)", TRUE),
                          downloadButton("download_poisson", "Download Results")
                        ),
                        mainPanel(
                          verbatimTextOutput("poisson_summary"),
                          DTOutput("poisson_table")
                        )
                      )
             )
           )
  ),
  
  # Survival Analysis Tab
  tabPanel("Survival Analysis", icon = icon("heartbeat"),
           tabsetPanel(
             tabPanel("Kaplan-Meier",
                      sidebarLayout(
                        sidebarPanel(
                          selectInput("surv_time", "Time Variable:", choices = NULL),
                          selectInput("surv_event", "Event Variable:", choices = NULL),
                          selectInput("surv_group", "Grouping Variable (Optional):", choices = NULL),
                          checkboxInput("surv_median", "Show Median Survival", TRUE),
                          checkboxInput("surv_mean", "Show Mean Survival", FALSE),
                          checkboxInput("surv_hazard", "Show Hazard Ratio", FALSE),
                          checkboxInput("surv_ph", "Test Proportional Hazards", FALSE),
                          colourInput("surv_color1", "Group 1 Color:", value = "#0C5EA8"),
                          colourInput("surv_color2", "Group 2 Color:", value = "#CD2026"),
                          textInput("surv_title", "Plot Title:", value = "Kaplan-Meier Curve"),
                          textInput("surv_xlab", "X-axis Label:", value = "Time"),
                          textInput("surv_ylab", "Y-axis Label:", value = "Survival Probability"),
                          textInput("surv_legend_title", "Legend Title:", value = "Group"),
                          textInput("surv_legend_labels", "Legend Labels (comma-separated):", value = ""),
                          numericInput("surv_break_time", "Time Break Interval:", value = NULL),
                          numericInput("surv_title_size", "Title Size:", value = 16),
                          numericInput("surv_x_size", "X-axis Label Size:", value = 14),
                          numericInput("surv_y_size", "Y-axis Label Size:", value = 14),
                          downloadButton("download_surv", "Download Plot")
                        ),
                        mainPanel(
                          verbatimTextOutput("surv_summary"),
                          plotOutput("surv_plot", height = "600px")
                        )
                      )
             ),
             tabPanel("Log-Rank Test",
                      sidebarLayout(
                        sidebarPanel(
                          selectInput("lr_time", "Time Variable:", choices = NULL),
                          selectInput("lr_event", "Event Variable:", choices = NULL),
                          selectInput("lr_group", "Grouping Variable:", choices = NULL),
                          downloadButton("download_logrank", "Download Results")
                        ),
                        mainPanel(
                          verbatimTextOutput("logrank_results")
                        )
                      )
             ),
             tabPanel("Cox Regression",
                      sidebarLayout(
                        sidebarPanel(
                          selectInput("cox_time", "Time Variable:", choices = NULL),
                          selectInput("cox_event", "Event Variable:", choices = NULL),
                          selectizeInput("cox_vars", "Predictor Variables:", choices = NULL, multiple = TRUE),
                          uiOutput("cox_ref_ui"),
                          sliderInput("cox_conf", "Confidence Level:", min = 0.90, max = 0.99, value = 0.95, step = 0.01),
                          downloadButton("download_cox", "Download Results")
                        ),
                        mainPanel(
                          verbatimTextOutput("cox_summary"),
                          DTOutput("cox_table")
                        )
                      )
             )
           )
  ),
  
  # Regression Analysis Tab
  tabPanel("Regression Analysis", icon = icon("line-chart"),
           tabsetPanel(
             tabPanel("Linear Regression",
                      sidebarLayout(
                        sidebarPanel(
                          selectInput("linear_outcome", "Outcome Variable:", choices = NULL),
                          selectizeInput("linear_vars", "Predictor Variables:", choices = NULL, multiple = TRUE),
                          uiOutput("linear_ref_ui"),
                          checkboxInput("linear_std", "Show Standardized Coefficients", FALSE),
                          downloadButton("download_linear", "Download Results")
                        ),
                        mainPanel(
                          verbatimTextOutput("linear_summary"),
                          DTOutput("linear_table")
                        )
                      )
             ),
             tabPanel("Logistic Regression",
                      sidebarLayout(
                        sidebarPanel(
                          selectInput("logistic_outcome", "Outcome Variable:", choices = NULL),
                          selectInput("logistic_type", "Regression Type:",
                                      choices = c("binary" = "binary", "ordinal" = "ordinal", "multinomial" = "multinomial"),
                                      selected = "binary"),
                          selectizeInput("logistic_vars", "Predictor Variables:", choices = NULL, multiple = TRUE),
                          uiOutput("logistic_ref_ui"),
                          checkboxInput("logistic_or", "Show Odds Ratios", TRUE),
                          downloadButton("download_logistic", "Download Results")
                        ),
                        mainPanel(
                          verbatimTextOutput("logistic_summary"),
                          DTOutput("logistic_table")
                        )
                      )
             )
           )
  ),
  
  # Statistical Tests Tab
  tabPanel("Statistical Tests", icon = icon("check-square"),
           tabsetPanel(
             tabPanel("T-tests",
                      sidebarLayout(
                        sidebarPanel(
                          selectInput("ttest_var", "Variable to Test:", choices = NULL),
                          selectInput("ttest_type", "Test Type:",
                                      choices = c("One Sample" = "one.sample",
                                                  "Two Sample" = "two.sample",
                                                  "Paired" = "paired",
                                                  "Mann-Whitney" = "mann.whitney"),
                                      selected = "two.sample"),
                          conditionalPanel(
                            condition = "input.ttest_type == 'one.sample'",
                            numericInput("ttest_mu", "Null Hypothesis Value:", value = 0)
                          ),
                          conditionalPanel(
                            condition = "input.ttest_type == 'two.sample' || input.ttest_type == 'paired' || input.ttest_type == 'mann.whitney'",
                            selectInput("ttest_group", "Grouping Variable:", choices = NULL)
                          ),
                          conditionalPanel(
                            condition = "input.ttest_type == 'two.sample'",
                            checkboxInput("ttest_var_equal", "Assume Equal Variances", FALSE)
                          ),
                          downloadButton("download_ttest", "Download Results")
                        ),
                        mainPanel(
                          verbatimTextOutput("ttest_results")
                        )
                      )
             ),
             tabPanel("ANOVA",
                      sidebarLayout(
                        sidebarPanel(
                          selectInput("anova_var", "Variable to Test:", choices = NULL),
                          selectInput("anova_group", "Grouping Variable:", choices = NULL),
                          selectInput("anova_type", "Test Type:",
                                      choices = c("ANOVA" = "anova", "Kruskal-Wallis" = "kruskal"),
                                      selected = "anova"),
                          checkboxInput("anova_posthoc", "Perform Post-hoc Tests", FALSE),
                          conditionalPanel(
                            condition = "input.anova_posthoc",
                            selectInput("anova_posthoc_type", "Post-hoc Test:",
                                        choices = c("Tukey HSD" = "tukey", "Dunn's Test" = "dunn"),
                                        selected = "tukey")
                          ),
                          downloadButton("download_anova", "Download Results")
                        ),
                        mainPanel(
                          verbatimTextOutput("anova_results"),
                          conditionalPanel(
                            condition = "input.anova_posthoc",
                            verbatimTextOutput("posthoc_results")
                          )
                        )
                      )
             ),
             tabPanel("Chi-square Test",
                      sidebarLayout(
                        sidebarPanel(
                          selectInput("chisq_var1", "Variable 1:", choices = NULL),
                          selectInput("chisq_var2", "Variable 2:", choices = NULL),
                          checkboxInput("chisq_expected", "Show Expected Counts", FALSE),
                          checkboxInput("chisq_residuals", "Show Residuals", FALSE),
                          downloadButton("download_chisq", "Download Results")
                        ),
                        mainPanel(
                          verbatimTextOutput("chisq_results"),
                          DTOutput("chisq_table")
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
    updateSelectInput(session, "surv_time", choices = names(df)[sapply(df, is.numeric)])
    updateSelectInput(session, "surv_event", choices = names(df))
    updateSelectInput(session, "surv_group", choices = c("None", names(df)))
    updateSelectInput(session, "lr_time", choices = names(df)[sapply(df, is.numeric)])
    updateSelectInput(session, "lr_event", choices = names(df))
    updateSelectInput(session, "lr_group", choices = names(df))
    updateSelectInput(session, "cox_time", choices = names(df)[sapply(df, is.numeric)])
    updateSelectInput(session, "cox_event", choices = names(df))
    updateSelectInput(session, "cox_vars", choices = names(df))
    updateSelectInput(session, "linear_outcome", choices = names(df)[sapply(df, is.numeric)])
    updateSelectInput(session, "linear_vars", choices = names(df))
    updateSelectInput(session, "logistic_outcome", choices = names(df))
    updateSelectInput(session, "logistic_vars", choices = names(df))
    updateSelectInput(session, "ttest_var", choices = names(df)[sapply(df, is.numeric)])
    updateSelectInput(session, "ttest_group", choices = names(df))
    updateSelectInput(session, "anova_var", choices = names(df)[sapply(df, is.numeric)])
    updateSelectInput(session, "anova_group", choices = names(df))
    updateSelectInput(session, "chisq_var1", choices = names(df))
    updateSelectInput(session, "chisq_var2", choices = names(df))
    
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
        
        ft <- flextable::flextable(freq_table)
        doc <- read_docx() %>% 
          officer::body_add_flextable(ft)
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
        
        ft <- flextable::flextable(central_table)
        doc <- read_docx() %>% 
          officer::body_add_flextable(ft)
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
  
  # Epidemic curve
  output$epi_curve <- renderPlotly({
    req(input$epi_date, input$epi_case)
    df <- data()
    
    # Convert date variable to Date class if it's character
    if(is.character(df[[input$epi_date]])) {
      df[[input$epi_date]] <- as.Date(df[[input$epi_date]])
    }
    
    # Aggregate data by time interval
    if(input$epi_interval == "Day") {
      df$time_group <- as.Date(df[[input$epi_date]])
    } else if(input$epi_interval == "Week") {
      df$time_group <- cut(df[[input$epi_date]], breaks = "week")
    } else if(input$epi_interval == "Month") {
      df$time_group <- cut(df[[input$epi_date]], breaks = "month")
    } else {
      df$time_group <- cut(df[[input$epi_date]], breaks = "year")
    }
    
    # Group data
    if(input$epi_group != "None") {
      epi_data <- df %>%
        group_by(time_group, .data[[input$epi_group]]) %>%
        summarise(cases = sum(.data[[input$epi_case]]), .groups = "drop")
      
      p <- ggplot(epi_data, aes(x = time_group, y = cases, fill = .data[[input$epi_group]])) +
        geom_col() +
        scale_fill_brewer(palette = "Set1")
    } else {
      epi_data <- df %>%
        group_by(time_group) %>%
        summarise(cases = sum(.data[[input$epi_case]]), .groups = "drop")
      
      p <- ggplot(epi_data, aes(x = time_group, y = cases)) +
        geom_col(fill = input$epi_color)
    }
    
    p <- p +
      labs(title = input$epi_title, x = "Date", y = "Number of Cases") +
      theme_minimal() +
      theme(
        plot.title = element_text(size = input$epi_title_size),
        axis.title.x = element_text(size = input$epi_x_size),
        axis.title.y = element_text(size = input$epi_y_size),
        axis.text.x = element_text(angle = 45, hjust = 1)
      )
    
    ggplotly(p)
  })
  
  output$download_epicurve <- downloadHandler(
    filename = function() {
      "epidemic_curve.png"
    },
    content = function(file) {
      df <- data()
      
      # Convert date variable to Date class if it's character
      if(is.character(df[[input$epi_date]])) {
        df[[input$epi_date]] <- as.Date(df[[input$epi_date]])
      }
      
      # Aggregate data by time interval
      if(input$epi_interval == "Day") {
        df$time_group <- as.Date(df[[input$epi_date]])
      } else if(input$epi_interval == "Week") {
        df$time_group <- cut(df[[input$epi_date]], breaks = "week")
      } else if(input$epi_interval == "Month") {
        df$time_group <- cut(df[[input$epi_date]], breaks = "month")
      } else {
        df$time_group <- cut(df[[input$epi_date]], breaks = "year")
      }
      
      # Group data
      if(input$epi_group != "None") {
        epi_data <- df %>%
          group_by(time_group, .data[[input$epi_group]]) %>%
          summarise(cases = sum(.data[[input$epi_case]]), .groups = "drop")
        
        p <- ggplot(epi_data, aes(x = time_group, y = cases, fill = .data[[input$epi_group]])) +
          geom_col() +
          scale_fill_brewer(palette = "Set1")
      } else {
        epi_data <- df %>%
          group_by(time_group) %>%
          summarise(cases = sum(.data[[input$epi_case]]), .groups = "drop")
        
        p <- ggplot(epi_data, aes(x = time_group, y = cases)) +
          geom_col(fill = input$epi_color)
      }
      
      p <- p +
        labs(title = input$epi_title, x = "Date", y = "Number of Cases") +
        theme_minimal() +
        theme(
          plot.title = element_text(size = input$epi_title_size),
          axis.title.x = element_text(size = input$epi_x_size),
          axis.title.y = element_text(size = input$epi_y_size),
          axis.text.x = element_text(angle = 45, hjust = 1)
        )
      
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
      
      ft <- flextable::flextable(demo_table)
      doc <- read_docx() %>% 
        officer::body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Incidence/Prevalence
  output$ip_table <- renderDT({
    req(input$ip_case, input$ip_pop)
    df <- data()
    
    if(input$ip_measure == "Incidence Rate") {
      req(input$ip_time)
      total_cases <- sum(df[[input$ip_case]], na.rm = TRUE)
      total_pop <- sum(df[[input$ip_pop]], na.rm = TRUE)
      total_time <- sum(df[[input$ip_time]], na.rm = TRUE)
      
      rate <- (total_cases / total_pop) * (1 / total_time) * 1000  # per 1000 person-time
      
      result <- data.frame(
        Measure = "Incidence Rate",
        Cases = total_cases,
        Population = total_pop,
        PersonTime = total_time,
        Rate = rate,
        Units = "per 1000 person-time"
      )
    } else if(input$ip_measure == "Cumulative Incidence") {
      total_cases <- sum(df[[input$ip_case]], na.rm = TRUE)
      total_pop <- sum(df[[input$ip_pop]], na.rm = TRUE)
      
      ci <- (total_cases / total_pop) * 1000  # per 1000 population
      
      result <- data.frame(
        Measure = "Cumulative Incidence",
        Cases = total_cases,
        Population = total_pop,
        Rate = ci,
        Units = "per 1000 population"
      )
    } else {
      # Prevalence
      total_cases <- sum(df[[input$ip_case]], na.rm = TRUE)
      total_pop <- sum(df[[input$ip_pop]], na.rm = TRUE)
      
      prevalence <- (total_cases / total_pop) * 100  # percentage
      
      result <- data.frame(
        Measure = "Prevalence",
        Cases = total_cases,
        Population = total_pop,
        Rate = prevalence,
        Units = "percent"
      )
    }
    
    datatable(result)
  })
  
  output$download_ip <- downloadHandler(
    filename = function() {
      "incidence_prevalence.docx"
    },
    content = function(file) {
      df <- data()
      
      if(input$ip_measure == "Incidence Rate") {
        req(input$ip_time)
        total_cases <- sum(df[[input$ip_case]], na.rm = TRUE)
        total_pop <- sum(df[[input$ip_pop]], na.rm = TRUE)
        total_time <- sum(df[[input$ip_time]], na.rm = TRUE)
        
        rate <- (total_cases / total_pop) * (1 / total_time) * 1000
        
        result <- data.frame(
          Measure = "Incidence Rate",
          Cases = total_cases,
          Population = total_pop,
          PersonTime = total_time,
          Rate = rate,
          Units = "per 1000 person-time"
        )
      } else if(input$ip_measure == "Cumulative Incidence") {
        total_cases <- sum(df[[input$ip_case]], na.rm = TRUE)
        total_pop <- sum(df[[input$ip_pop]], na.rm = TRUE)
        
        ci <- (total_cases / total_pop) * 1000
        
        result <- data.frame(
          Measure = "Cumulative Incidence",
          Cases = total_cases,
          Population = total_pop,
          Rate = ci,
          Units = "per 1000 population"
        )
      } else {
        total_cases <- sum(df[[input$ip_case]], na.rm = TRUE)
        total_pop <- sum(df[[input$ip_pop]], na.rm = TRUE)
        
        prevalence <- (total_cases / total_pop) * 100
        
        result <- data.frame(
          Measure = "Prevalence",
          Cases = total_cases,
          Population = total_pop,
          Rate = prevalence,
          Units = "percent"
        )
      }
      
      ft <- flextable::flextable(result)
      doc <- read_docx() %>% 
        officer::body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Attack Rates Table
  output$ar_table <- renderDT({
    req(input$ar_case, input$ar_pop)
    df <- data()
    
    if (input$ar_group != "None") {
      ar_data <- df %>%
        group_by(.data[[input$ar_group]]) %>%
        summarise(
          Cases = sum(.data[[input$ar_case]], na.rm = TRUE),
          Population = sum(.data[[input$ar_pop]], na.rm = TRUE),
          AttackRate = (Cases / Population) * 100,
          .groups = "drop"
        )
    } else {
      ar_data <- data.frame(
        Cases = sum(df[[input$ar_case]], na.rm = TRUE),
        Population = sum(df[[input$ar_pop]], na.rm = TRUE),
        AttackRate = (sum(df[[input$ar_case]], na.rm = TRUE) / sum(df[[input$ar_pop]], na.rm = TRUE)) * 100
      )
    }
    
    datatable(ar_data)
  })
  
  # Attack Rates with Confidence Intervals
  output$ar_detailed_table <- renderDT({
    req(input$ar_case, input$ar_pop)
    df <- data()
    
    if (input$ar_group != "None") {
      ar_data <- df %>%
        group_by(.data[[input$ar_group]]) %>%
        summarise(
          Cases = sum(.data[[input$ar_case]], na.rm = TRUE),
          Population = sum(.data[[input$ar_pop]], na.rm = TRUE),
          AttackRate = (Cases / Population) * 100,
          LowerCI = prop.test(Cases, Population)$conf.int[1] * 100,
          UpperCI = prop.test(Cases, Population)$conf.int[2] * 100,
          .groups = "drop"
        )
    } else {
      total_cases <- sum(df[[input$ar_case]], na.rm = TRUE)
      total_pop <- sum(df[[input$ar_pop]], na.rm = TRUE)
      prop_test <- prop.test(total_cases, total_pop)
      
      ar_data <- data.frame(
        Cases = total_cases,
        Population = total_pop,
        AttackRate = (total_cases / total_pop) * 100,
        LowerCI = prop_test$conf.int[1] * 100,
        UpperCI = prop_test$conf.int[2] * 100
      )
    }
    
    datatable(ar_data)
  })
  
  # Download Handler for Attack Rate Table
  output$download_ar <- downloadHandler(
    filename = function() {
      "attack_rates.docx"
    },
    content = function(file) {
      df <- data()
      
      if (input$ar_format == "Table") {
        if (input$ar_group != "None") {
          ar_data <- df %>%
            group_by(.data[[input$ar_group]]) %>%
            summarise(
              Cases = sum(.data[[input$ar_case]], na.rm = TRUE),
              Population = sum(.data[[input$ar_pop]], na.rm = TRUE),
              AttackRate = (Cases / Population) * 100,
              .groups = "drop"
            )
        } else {
          ar_data <- data.frame(
            Cases = sum(df[[input$ar_case]], na.rm = TRUE),
            Population = sum(df[[input$ar_pop]], na.rm = TRUE),
            AttackRate = (sum(df[[input$ar_case]], na.rm = TRUE) / sum(df[[input$ar_pop]], na.rm = TRUE)) * 100
          )
        }
      } else {
        if (input$ar_group != "None") {
          ar_data <- df %>%
            group_by(.data[[input$ar_group]]) %>%
            summarise(
              Cases = sum(.data[[input$ar_case]], na.rm = TRUE),
              Population = sum(.data[[input$ar_pop]], na.rm = TRUE),
              AttackRate = (Cases / Population) * 100,
              LowerCI = prop.test(Cases, Population)$conf.int[1] * 100,
              UpperCI = prop.test(Cases, Population)$conf.int[2] * 100,
              .groups = "drop"
            )
        } else {
          total_cases <- sum(df[[input$ar_case]], na.rm = TRUE)
          total_pop <- sum(df[[input$ar_pop]], na.rm = TRUE)
          prop_test <- prop.test(total_cases, total_pop)
          
          ar_data <- data.frame(
            Cases = total_cases,
            Population = total_pop,
            AttackRate = (total_cases / total_pop) * 100,
            LowerCI = prop_test$conf.int[1] * 100,
            UpperCI = prop_test$conf.int[2] * 100
          )
        }
      }
      
      ft <- flextable::flextable(ar_data)
      doc <- read_docx() %>%
        officer::body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  
  # Mortality Rates
  output$mort_table <- renderDT({
    req(input$mort_case, input$mort_pop)
    df <- data()
    
    if(input$mort_measure == "Mortality Rate") {
      total_cases <- sum(df[[input$mort_case]], na.rm = TRUE)
      total_pop <- sum(df[[input$mort_pop]], na.rm = TRUE)
      
      rate <- (total_cases / total_pop) * 1000  # per 1000 population
      
      result <- data.frame(
        Measure = "Mortality Rate",
        Deaths = total_cases,
        Population = total_pop,
        Rate = rate,
        Units = "per 1000 population"
      )
    } else {
      # Case Fatality Rate
      total_cases <- sum(df[[input$mort_case]], na.rm = TRUE)
      total_cases_all <- nrow(df)
      
      cfr <- (total_cases / total_cases_all) * 100  # percentage
      
      result <- data.frame(
        Measure = "Case Fatality Rate",
        Deaths = total_cases,
        Cases = total_cases_all,
        Rate = cfr,
        Units = "percent"
      )
    }
    
    datatable(result)
  })
  
  output$download_mort <- downloadHandler(
    filename = function() {
      "mortality_rates.docx"
    },
    content = function(file) {
      df <- data()
      
      if(input$mort_measure == "Mortality Rate") {
        total_cases <- sum(df[[input$mort_case]], na.rm = TRUE)
        total_pop <- sum(df[[input$mort_pop]], na.rm = TRUE)
        
        rate <- (total_cases / total_pop) * 1000
        
        result <- data.frame(
          Measure = "Mortality Rate",
          Deaths = total_cases,
          Population = total_pop,
          Rate = rate,
          Units = "per 1000 population"
        )
      } else {
        total_cases <- sum(df[[input$mort_case]], na.rm = TRUE)
        total_cases_all <- nrow(df)
        
        cfr <- (total_cases / total_cases_all) * 100
        
        result <- data.frame(
          Measure = "Case Fatality Rate",
          Deaths = total_cases,
          Cases = total_cases_all,
          Rate = cfr,
          Units = "percent"
        )
      }
      
      ft <- flextable::flextable(result)
      doc <- read_docx() %>% 
        officer::body_add_flextable(ft)
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
      
      ft <- flextable::flextable(lifetable) %>%
        flextable::set_caption(input$lt_title) %>%
        flextable::theme_zebra()
      
      doc <- read_docx() %>% 
        officer::body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Risk Ratios
  output$rr_results <- renderPrint({
    req(input$rr_outcome, input$rr_exposure)
    df <- data()
    
    # Create 2x2 table
    tab <- table(df[[input$rr_exposure]], df[[input$rr_outcome]])
    
    # Calculate risk ratio with confidence intervals
    rr <- epitools::riskratio(tab, conf.level = input$rr_conf)
    
    print(rr)
  })
  
  output$rr_table <- renderDT({
    req(input$rr_outcome, input$rr_exposure)
    df <- data()
    
    tab <- table(df[[input$rr_exposure]], df[[input$rr_outcome]])
    rr <- epitools::riskratio(tab, conf.level = input$rr_conf)
    
    # Create summary table
    rr_table <- data.frame(
      Measure = c("Risk Ratio", "Lower CI", "Upper CI"),
      Value = c(rr$measure[2,1], rr$measure[2,2], rr$measure[2,3])
    )
    
    datatable(rr_table)
  })
  
  output$download_rr <- downloadHandler(
    filename = function() {
      "risk_ratio.docx"
    },
    content = function(file) {
      df <- data()
      
      tab <- table(df[[input$rr_exposure]], df[[input$rr_outcome]])
      rr <- epitools::riskratio(tab, conf.level = input$rr_conf)
      
      rr_table <- data.frame(
        Measure = c("Risk Ratio", "Lower CI", "Upper CI"),
        Value = c(rr$measure[2,1], rr$measure[2,2], rr$measure[2,3])
      )
      
      ft <- flextable::flextable(rr_table)
      doc <- read_docx() %>% 
        officer::body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Odds Ratios
  output$or_results <- renderPrint({
    req(input$or_outcome, input$or_exposure)
    df <- data()
    
    # Create 2x2 table
    tab <- table(df[[input$or_exposure]], df[[input$or_outcome]])
    
    # Calculate odds ratio with confidence intervals
    or <- epitools::oddsratio(tab, conf.level = input$or_conf)
    
    print(or)
  })
  
  output$or_table <- renderDT({
    req(input$or_outcome, input$or_exposure)
    df <- data()
    
    tab <- table(df[[input$or_exposure]], df[[input$or_outcome]])
    or <- epitools::oddsratio(tab, conf.level = input$or_conf)
    
    # Create summary table
    or_table <- data.frame(
      Measure = c("Odds Ratio", "Lower CI", "Upper CI"),
      Value = c(or$measure[2,1], or$measure[2,2], or$measure[2,3])
    )
    
    datatable(or_table)
  })
  
  output$download_or <- downloadHandler(
    filename = function() {
      "odds_ratio.docx"
    },
    content = function(file) {
      df <- data()
      
      tab <- table(df[[input$or_exposure]], df[[input$or_outcome]])
      or <- epitools::oddsratio(tab, conf.level = input$or_conf)
      
      or_table <- data.frame(
        Measure = c("Odds Ratio", "Lower CI", "Upper CI"),
        Value = c(or$measure[2,1], or$measure[2,2], or$measure[2,3])
      )
      
      ft <- flextable::flextable(or_table)
      doc <- read_docx() %>% 
        officer::body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Attributable Risk
  output$ar_results <- renderPrint({
    req(input$ar_outcome, input$ar_exposure)
    df <- data()
    
    # Create 2x2 table
    tab <- table(df[[input$ar_exposure]], df[[input$ar_outcome]])
    
    # Calculate attributable risk
    ar <- epitools::riskratio(tab, conf.level = input$ar_conf)
    
    print(ar)
  })
  
  output$ar_table <- renderDT({
    req(input$ar_outcome, input$ar_exposure)
    df <- data()
    
    tab <- table(df[[input$ar_exposure]], df[[input$ar_outcome]])
    ar <- epitools::riskratio(tab, conf.level = input$ar_conf)
    
    # Create summary table
    ar_table <- data.frame(
      Measure = c("Attributable Risk", "Lower CI", "Upper CI"),
      Value = c(ar$measure[1,1], ar$measure[1,2], ar$measure[1,3])
    )
    
    if(input$ar_paf) {
      # Calculate population attributable fraction
      paf <- (ar$measure[1,1] - 1) / ar$measure[1,1]
      ar_table <- rbind(ar_table, data.frame(Measure = "Population Attributable Fraction", Value = paf))
    }
    
    datatable(ar_table)
  })
  
  output$download_arisk <- downloadHandler(
    filename = function() {
      "attributable_risk.docx"
    },
    content = function(file) {
      df <- data()
      
      tab <- table(df[[input$ar_exposure]], df[[input$ar_outcome]])
      ar <- epitools::riskratio(tab, conf.level = input$ar_conf)
      
      ar_table <- data.frame(
        Measure = c("Attributable Risk", "Lower CI", "Upper CI"),
        Value = c(ar$measure[1,1], ar$measure[1,2], ar$measure[1,3])
      )
      
      if(input$ar_paf) {
        paf <- (ar$measure[1,1] - 1) / ar$measure[1,1]
        ar_table <- rbind(ar_table, data.frame(Measure = "Population Attributable Fraction", Value = paf))
      }
      
      ft <- flextable::flextable(ar_table)
      doc <- read_docx() %>% 
        officer::body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Poisson Regression Reference UI
  output$poisson_ref_ui <- renderUI({
    req(input$poisson_vars)
    df <- data()
    
    ref_inputs <- lapply(input$poisson_vars, function(var) {
      if (is.factor(df[[var]])) {
        selectInput(paste0("poisson_ref_", var),
                    paste("Reference level for", var),
                    choices = levels(df[[var]]))
      } else if (is.character(df[[var]])) {
        selectInput(paste0("poisson_ref_", var),
                    paste("Reference level for", var),
                    choices = unique(df[[var]]))
      }
    })
    
    do.call(tagList, ref_inputs)
  })
  
  # Poisson Regression Summary Output
  output$poisson_summary <- renderPrint({
    req(input$poisson_outcome, input$poisson_vars)
    df <- data()
    
    # Set reference levels
    for (var in input$poisson_vars) {
      ref_input <- input[[paste0("poisson_ref_", var)]]
      if (!is.null(ref_input) && (is.factor(df[[var]]) || is.character(df[[var]]))) {
        df[[var]] <- relevel(as.factor(df[[var]]), ref = ref_input)
      }
    }
    
    # Create formula
    formula <- as.formula(paste(input$poisson_outcome, "~", paste(input$poisson_vars, collapse = "+")))
    
    # Fit model
    if (!is.null(input$poisson_offset) && input$poisson_offset != "") {
      model <- glm(formula, family = poisson(link = "log"),
                   data = df, offset = log(df[[input$poisson_offset]]))
    } else {
      model <- glm(formula, family = poisson(link = "log"), data = df)
    }
    
    # Print summary
    summary(model)
  })
  
  # Poisson Regression Table Output
  output$poisson_table <- renderDT({
    req(input$poisson_outcome, input$poisson_vars)
    df <- data()
    
    # Set reference levels
    for (var in input$poisson_vars) {
      ref_input <- input[[paste0("poisson_ref_", var)]]
      if (!is.null(ref_input) && (is.factor(df[[var]]) || is.character(df[[var]]))) {
        df[[var]] <- relevel(as.factor(df[[var]]), ref = ref_input)
      }
    }
    
    # Create formula
    formula <- as.formula(paste(input$poisson_outcome, "~", paste(input$poisson_vars, collapse = "+")))
    
    # Fit model
    if (!is.null(input$poisson_offset) && input$poisson_offset != "") {
      model <- glm(formula, family = poisson(link = "log"),
                   data = df, offset = log(df[[input$poisson_offset]]))
    } else {
      model <- glm(formula, family = poisson(link = "log"), data = df)
    }
    
    # Create tidy table
    if (input$poisson_irr) {
      poisson_table <- broom::tidy(model, exponentiate = TRUE, conf.int = TRUE, conf.level = input$poisson_conf)
    } else {
      poisson_table <- broom::tidy(model, conf.int = TRUE, conf.level = input$poisson_conf)
    }
    
    datatable(poisson_table)
  })
  
  # Poisson Regression Download
  output$download_poisson <- downloadHandler(
    filename = function() {
      "poisson_regression.docx"
    },
    content = function(file) {
      df <- data()
      
      # Set reference levels
      for (var in input$poisson_vars) {
        ref_input <- input[[paste0("poisson_ref_", var)]]
        if (!is.null(ref_input) && (is.factor(df[[var]]) || is.character(df[[var]]))) {
          df[[var]] <- relevel(as.factor(df[[var]]), ref = ref_input)
        }
      }
      
      # Create formula
      formula <- as.formula(paste(input$poisson_outcome, "~", paste(input$poisson_vars, collapse = "+")))
      
      # Fit model
      if (!is.null(input$poisson_offset) && input$poisson_offset != "") {
        model <- glm(formula, family = poisson(link = "log"),
                     data = df, offset = log(df[[input$poisson_offset]]))
      } else {
        model <- glm(formula, family = poisson(link = "log"), data = df)
      }
      
      # Create tidy table
      if (input$poisson_irr) {
        poisson_table <- broom::tidy(model, exponentiate = TRUE, conf.int = TRUE, conf.level = input$poisson_conf)
      } else {
        poisson_table <- broom::tidy(model, conf.int = TRUE, conf.level = input$poisson_conf)
      }
      
      ft <- flextable::flextable(poisson_table)
      doc <- officer::read_docx() %>%
        officer::body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Survival Analysis (Kaplan-Meier)
  output$surv_summary <- renderPrint({
    req(input$surv_time, input$surv_event)
    df <- data()
    
    # Create survival object
    surv_obj <- create_surv_obj(input$surv_time, input$surv_event, df)
    
    # Fit survival model
    if(input$surv_group != "None") {
      surv_fit <- survfit(surv_obj ~ .data[[input$surv_group]], data = df)
    } else {
      surv_fit <- survfit(surv_obj ~ 1, data = df)
    }
    
    # Print summary
    print(surv_fit)
    
    # Show median survival if requested
    if(input$surv_median) {
      cat("\nMedian Survival:\n")
      print(survminer::surv_median(surv_fit))
    }
    
    # Show mean survival if requested
    if(input$surv_mean) {
      cat("\nMean Survival:\n")
      mean_surv <- survfit(surv_obj ~ 1)
      print(mean(mean_surv$time))
    }
    
    # Show hazard ratio if requested and groups exist
    if(input$surv_hazard && input$surv_group != "None") {
      cat("\nHazard Ratio:\n")
      cox_model <- coxph(surv_obj ~ .data[[input$surv_group]], data = df)
      print(summary(cox_model))
    }
    
    # Test proportional hazards if requested
    if(input$surv_ph && input$surv_group != "None") {
      cat("\nProportional Hazards Test:\n")
      cox_model <- coxph(surv_obj ~ .data[[input$surv_group]], data = df)
      print(cox.zph(cox_model))
    }
  })
  
  output$surv_plot <- renderPlot({
    req(input$surv_time, input$surv_event)
    df <- data()
    
    # Create survival object
    surv_obj <- create_surv_obj(input$surv_time, input$surv_event, df)
    
    # Fit survival model
    if(input$surv_group != "None") {
      surv_fit <- survfit(surv_obj ~ .data[[input$surv_group]], data = df)
      
      # Customize legend labels if provided
      legend_labels <- unlist(strsplit(input$surv_legend_labels, ","))
      if(length(legend_labels) > 0 && length(legend_labels) == length(levels(as.factor(df[[input$surv_group]])))) {
        names(legend_labels) <- levels(as.factor(df[[input$surv_group]]))
      } else {
        legend_labels <- NULL
      }
    } else {
      surv_fit <- survfit(surv_obj ~ 1, data = df)
      legend_labels <- NULL
    }
    
    # Create plot
    p <- survminer::ggsurvplot(
      surv_fit,
      data = df,
      risk.table = TRUE,
      pval = TRUE,
      conf.int = TRUE,
      palette = c(input$surv_color1, input$surv_color2),
      title = input$surv_title,
      xlab = input$surv_xlab,
      ylab = input$surv_ylab,
      legend.title = input$surv_legend_title,
      legend.labs = legend_labels,
      break.time.by = input$surv_break_time,
      ggtheme = theme_minimal(),
      font.title = c(input$surv_title_size, "bold"),
      font.x = c(input$surv_x_size, "plain"),
      font.y = c(input$surv_y_size, "plain"),
      font.tickslab = c(12, "plain")
    )
    
    print(p)
  })
  
  output$download_surv <- downloadHandler(
    filename = function() {
      "survival_plot.png"
    },
    content = function(file) {
      df <- data()
      
      surv_obj <- create_surv_obj(input$surv_time, input$surv_event, df)
      
      if(input$surv_group != "None") {
        surv_fit <- survfit(surv_obj ~ .data[[input$surv_group]], data = df)
        
        legend_labels <- unlist(strsplit(input$surv_legend_labels, ","))
        if(length(legend_labels) > 0 && length(legend_labels) == length(levels(as.factor(df[[input$surv_group]])))) {
          names(legend_labels) <- levels(as.factor(df[[input$surv_group]]))
        } else {
          legend_labels <- NULL
        }
      } else {
        surv_fit <- survfit(surv_obj ~ 1, data = df)
        legend_labels <- NULL
      }
      
      p <- survminer::ggsurvplot(
        surv_fit,
        data = df,
        risk.table = TRUE,
        pval = TRUE,
        conf.int = TRUE,
        palette = c(input$surv_color1, input$surv_color2),
        title = input$surv_title,
        xlab = input$surv_xlab,
        ylab = input$surv_ylab,
        legend.title = input$surv_legend_title,
        legend.labs = legend_labels,
        break.time.by = input$surv_break_time,
        ggtheme = theme_minimal(),
        font.title = c(input$surv_title_size, "bold"),
        font.x = c(input$surv_x_size, "plain"),
        font.y = c(input$surv_y_size, "plain"),
        font.tickslab = c(12, "plain")
      )
      
      ggsave(file, plot = print(p), device = "png", width = 10, height = 8)
    }
  )
  
  # Log-rank test
  output$logrank_results <- renderPrint({
    req(input$lr_time, input$lr_event, input$lr_group)
    df <- data()
    
    surv_obj <- create_surv_obj(input$lr_time, input$lr_event, df)
    
    fit <- survdiff(surv_obj ~ df[[input$lr_group]])
    
    cat("Log-Rank Test Results:\n")
    print(fit)
  })
  
  output$download_logrank <- downloadHandler(
    filename = "logrank_test.docx",
    content = function(file) {
      df <- data()
      
      surv_obj <- create_surv_obj(input$lr_time, input$lr_event, df)
      
      fit <- survdiff(surv_obj ~ df[[input$lr_group]])
      
      # Convert to flextable
      res <- capture.output(fit)
      ft <- flextable::flextable(data.frame(Results = res))
      doc <- read_docx() %>% 
        officer::body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  
  # Cox Regression
  output$cox_ref_ui <- renderUI({
    req(input$cox_vars)
    df <- data()
    
    ref_inputs <- lapply(input$cox_vars, function(var) {
      if(is.factor(df[[var]])) {
        selectInput(paste0("cox_ref_", var), 
                    paste("Reference level for", var), 
                    choices = levels(df[[var]]))
      } else if(is.character(df[[var]])) {
        selectInput(paste0("cox_ref_", var), 
                    paste("Reference level for", var), 
                    choices = unique(df[[var]]))
      }
    })
    
    do.call(tagList, ref_inputs)
  })
  
  output$cox_summary <- renderPrint({
    req(input$cox_time, input$cox_event, input$cox_vars)
    df <- data()
    
    # Set reference levels
    for(var in input$cox_vars) {
      ref_input <- input[[paste0("cox_ref_", var)]]
      if(!is.null(ref_input) && (is.factor(df[[var]]) || is.character(df[[var]]))) {
        df[[var]] <- relevel(as.factor(df[[var]]), ref = ref_input)
      }
    }
    
    # Create survival object
    surv_obj <- create_surv_obj(input$cox_time, input$cox_event, df)
    
    # Create formula
    formula <- as.formula(paste("surv_obj ~", paste(input$cox_vars, collapse = "+")))
    
    # Fit Cox model
    model <- coxph(formula, data = df)
    
    # Print summary
    summary(model)
  })
  
  output$cox_table <- renderDT({
    req(input$cox_time, input$cox_event, input$cox_vars)
    df <- data()
    
    # Set reference levels
    for(var in input$cox_vars) {
      ref_input <- input[[paste0("cox_ref_", var)]]
      if(!is.null(ref_input) && (is.factor(df[[var]]) || is.character(df[[var]]))) {
        df[[var]] <- relevel(as.factor(df[[var]]), ref = ref_input)
      }
    }
    
    # Create survival object
    surv_obj <- create_surv_obj(input$cox_time, input$cox_event, df)
    
    # Create formula
    formula <- as.formula(paste("surv_obj ~", paste(input$cox_vars, collapse = "+")))
    
    # Fit Cox model
    model <- coxph(formula, data = df)
    
    # Create tidy table
    cox_table <- broom::tidy(model, conf.int = TRUE, conf.level = input$cox_conf, exponentiate = TRUE)
    
    datatable(cox_table)
  })
  
  output$download_cox <- downloadHandler(
    filename = function() {
      "cox_regression.docx"
    },
    content = function(file) {
      df <- data()
      
      # Set reference levels
      for(var in input$cox_vars) {
        ref_input <- input[[paste0("cox_ref_", var)]]
        if(!is.null(ref_input) && (is.factor(df[[var]]) || is.character(df[[var]]))) {
          df[[var]] <- relevel(as.factor(df[[var]]), ref = ref_input)
        }
      }
      
      surv_obj <- create_surv_obj(input$cox_time, input$cox_event, df)
      formula <- as.formula(paste("surv_obj ~", paste(input$cox_vars, collapse = "+")))
      model <- coxph(formula, data = df)
      
      cox_table <- broom::tidy(model, conf.int = TRUE, conf.level = input$cox_conf, exponentiate = TRUE)
      
      ft <- flextable::flextable(cox_table)
      doc <- read_docx() %>% 
        officer::body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Linear regression with reference category selection
  output$linear_ref_ui <- renderUI({
    req(input$linear_vars)
    df <- data()
    
    ref_ui <- tagList()
    
    for (var in input$linear_vars) {
      if (is.factor(df[[var]]) || is.character(df[[var]])) {
        levels <- if (is.factor(df[[var]])) levels(df[[var]]) else unique(df[[var]])
        ref_ui <- tagAppendChild(
          ref_ui,
          selectInput(paste0("linear_ref_", var), 
                      paste("Reference category for", var), 
                      choices = levels)
        )
      }
    }
    
    ref_ui
  })
  
  output$linear_summary <- renderPrint({
    req(input$linear_outcome, input$linear_vars)
    df <- data()
    
    # Set reference categories for factors
    for (var in input$linear_vars) {
      if (is.factor(df[[var]])) {
        ref <- input[[paste0("linear_ref_", var)]]
        if (!is.null(ref)) {
          df[[var]] <- relevel(df[[var]], ref = ref)
        }
      } else if (is.character(df[[var]])) {
        ref <- input[[paste0("linear_ref_", var)]]
        if (!is.null(ref)) {
          df[[var]] <- factor(df[[var]])
          df[[var]] <- relevel(df[[var]], ref = ref)
        }
      }
    }
    
    # Create formula
    formula <- as.formula(paste(input$linear_outcome, "~", paste(input$linear_vars, collapse = " + ")))
    
    # Fit linear model
    linear_fit <- lm(formula, data = df)
    
    cat("Linear Regression Results:\n\n")
    print(summary(linear_fit))
    
    if (input$linear_std) {
      cat("\nStandardized Coefficients:\n")
      # Calculate standardized coefficients
      std_coef <- data.frame(
        term = names(coef(linear_fit)),
        estimate = lm.beta::lm.beta(linear_fit)$standardized.coefficients
      )
      print(std_coef)
    }
  })
  
  output$linear_table <- renderDT({
    req(input$linear_outcome, input$linear_vars)
    df <- data()
    
    # Set reference categories for factors
    for (var in input$linear_vars) {
      if (is.factor(df[[var]])) {
        ref <- input[[paste0("linear_ref_", var)]]
        if (!is.null(ref)) {
          df[[var]] <- relevel(df[[var]], ref = ref)
        }
      } else if (is.character(df[[var]])) {
        ref <- input[[paste0("linear_ref_", var)]]
        if (!is.null(ref)) {
          df[[var]] <- factor(df[[var]])
          df[[var]] <- relevel(df[[var]], ref = ref)
        }
      }
    }
    
    # Create formula
    formula <- as.formula(paste(input$linear_outcome, "~", paste(input$linear_vars, collapse = " + ")))
    
    # Fit linear model
    linear_fit <- lm(formula, data = df)
    
    # Create results table
    results <- broom::tidy(linear_fit, conf.int = TRUE) %>%
      select(term, estimate, conf.low, conf.high, p.value) %>%
      rename(Coefficient = estimate, `Lower CI` = conf.low, `Upper CI` = conf.high, `p-value` = p.value)
    
    if (input$linear_std) {
      # Add standardized coefficients
      std_coef <- data.frame(
        term = names(coef(linear_fit)),
        std_estimate = lm.beta::lm.beta(linear_fit)$standardized.coefficients
      )
      results <- results %>%
        left_join(std_coef, by = "term") %>%
        rename(`Standardized Coefficient` = std_estimate)
    }
    
    datatable(results, options = list(pageLength = 10))
  })
  
  output$download_linear <- downloadHandler(
    filename = "linear_regression.docx",
    content = function(file) {
      df <- data()
      
      # Set reference categories for factors
      for (var in input$linear_vars) {
        if (is.factor(df[[var]]) || is.character(df[[var]])) {
          ref <- input[[paste0("linear_ref_", var)]]
          if (!is.null(ref)) {
            df[[var]] <- factor(df[[var]])
            df[[var]] <- relevel(df[[var]], ref = ref)
          }
        }
      }
      
      formula <- as.formula(paste(input$linear_outcome, "~", paste(input$linear_vars, collapse = " + ")))
      linear_fit <- lm(formula, data = df)
      
      # Create results table
      results <- broom::tidy(linear_fit, conf.int = TRUE) %>%
        select(term, estimate, conf.low, conf.high, p.value) %>%
        rename(Coefficient = estimate, `Lower CI` = conf.low, `Upper CI` = conf.high, `p-value` = p.value)
      
      if (input$linear_std) {
        # Add standardized coefficients
        std_coef <- data.frame(
          term = names(coef(linear_fit)),
          std_estimate = lm.beta::lm.beta(linear_fit)$standardized.coefficients
        )
        results <- results %>%
          left_join(std_coef, by = "term") %>%
          rename(`Standardized Coefficient` = std_estimate)
      }
      
      ft <- flextable::flextable(results) %>%
        flextable::set_caption("Linear Regression Results") %>%
        flextable::theme_zebra()
      
      doc <- read_docx() %>% 
        officer::body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Logistic Regression
  output$logistic_ref_ui <- renderUI({
    req(input$logistic_vars)
    df <- data()
    
    ref_inputs <- lapply(input$logistic_vars, function(var) {
      if(is.factor(df[[var]])) {
        selectInput(paste0("logistic_ref_", var), 
                    paste("Reference level for", var), 
                    choices = levels(df[[var]]))
      } else if(is.character(df[[var]])) {
        selectInput(paste0("logistic_ref_", var), 
                    paste("Reference level for", var), 
                    choices = unique(df[[var]]))
      }
    })
    
    do.call(tagList, ref_inputs)
  })
  
  output$logistic_summary <- renderPrint({
    req(input$logistic_outcome, input$logistic_vars)
    df <- data()
    
    # Set reference levels
    for(var in input$logistic_vars) {
      ref_input <- input[[paste0("logistic_ref_", var)]]
      if(!is.null(ref_input) && (is.factor(df[[var]]) || is.character(df[[var]]))) {
        df[[var]] <- relevel(as.factor(df[[var]]), ref = ref_input)
      }
    }
    
    # Create formula
    formula <- as.formula(paste(input$logistic_outcome, "~", paste(input$logistic_vars, collapse = "+")))
    
    # Fit appropriate model based on type
    if(input$logistic_type == "binary") {
      model <- glm(formula, family = binomial(link = "logit"), data = df)
    } else if(input$logistic_type == "ordinal") {
      model <- MASS::polr(formula, data = df, Hess = TRUE)
    } else {
      model <- nnet::multinom(formula, data = df)
    }
    
    # Print summary
    print(summary(model))
    
    # Show odds ratios for binary logistic regression
    if(input$logistic_or && input$logistic_type == "binary") {
      cat("\nOdds Ratios:\n")
      or <- exp(coef(model))
      ci <- exp(confint(model))
      or_table <- data.frame(
        OR = or,
        LowerCI = ci[,1],
        UpperCI = ci[,2]
      )
      print(or_table)
    }
  })
  
  output$logistic_table <- renderDT({
    req(input$logistic_outcome, input$logistic_vars)
    df <- data()
    
    # Set reference levels
    for(var in input$logistic_vars) {
      ref_input <- input[[paste0("logistic_ref_", var)]]
      if(!is.null(ref_input) && (is.factor(df[[var]]) || is.character(df[[var]]))) {
        df[[var]] <- relevel(as.factor(df[[var]]), ref = ref_input)
      }
    }
    
    # Create formula
    formula <- as.formula(paste(input$logistic_outcome, "~", paste(input$logistic_vars, collapse = "+")))
    
    # Fit appropriate model
    if(input$logistic_type == "binary") {
      model <- glm(formula, family = binomial(link = "logit"), data = df)
      
      # Create tidy table
      if(input$logistic_or) {
        logistic_table <- broom::tidy(model, conf.int = TRUE, exponentiate = TRUE)
      } else {
        logistic_table <- broom::tidy(model, conf.int = TRUE)
      }
    } else if(input$logistic_type == "ordinal") {
      model <- MASS::polr(formula, data = df, Hess = TRUE)
      logistic_table <- as.data.frame(coef(summary(model)))
      logistic_table$p.value <- pnorm(abs(logistic_table[, "t value"]), lower.tail = FALSE) * 2
    } else {
      model <- nnet::multinom(formula, data = df)
      logistic_table <- as.data.frame(coef(summary(model)))
    }
    
    datatable(logistic_table)
  })
  
  output$download_logistic <- downloadHandler(
    filename = function() {
      "logistic_regression.docx"
    },
    content = function(file) {
      df <- data()
      
      # Set reference levels
      for(var in input$logistic_vars) {
        ref_input <- input[[paste0("logistic_ref_", var)]]
        if(!is.null(ref_input) && (is.factor(df[[var]]) || is.character(df[[var]]))) {
          df[[var]] <- relevel(as.factor(df[[var]]), ref = ref_input)
        }
      }
      
      formula <- as.formula(paste(input$logistic_outcome, "~", paste(input$logistic_vars, collapse = "+")))
      
      if(input$logistic_type == "binary") {
        model <- glm(formula, family = binomial(link = "logit"), data = df)
        
        if(input$logistic_or) {
          logistic_table <- broom::tidy(model, conf.int = TRUE, exponentiate = TRUE)
        } else {
          logistic_table <- broom::tidy(model, conf.int = TRUE)
        }
      } else if(input$logistic_type == "ordinal") {
        model <- MASS::polr(formula, data = df, Hess = TRUE)
        logistic_table <- as.data.frame(coef(summary(model)))
        logistic_table$p.value <- pnorm(abs(logistic_table[, "t value"]), lower.tail = FALSE) * 2
      } else {
        model <- nnet::multinom(formula, data = df)
        logistic_table <- as.data.frame(coef(summary(model)))
      }
      
      ft <- flextable::flextable(logistic_table)
      doc <- read_docx() %>% 
        officer::body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # T-tests/Mann-Whitney
  output$ttest_results <- renderPrint({
    req(input$ttest_var)
    df <- data()
    
    if(input$ttest_type == "one.sample") {
      # One sample t-test
      t_test <- t.test(df[[input$ttest_var]], mu = input$ttest_mu)
      print(t_test)
    } else if(input$ttest_type == "two.sample") {
      # Two sample t-test
      req(input$ttest_group)
      t_test <- t.test(df[[input$ttest_var]] ~ df[[input$ttest_group]], 
                       var.equal = input$ttest_var_equal)
      print(t_test)
    } else if(input$ttest_type == "paired") {
      # Paired t-test
      req(input$ttest_group)
      t_test <- t.test(df[[input$ttest_var]], df[[input$ttest_group]], paired = TRUE)
      print(t_test)
    } else {
      # Mann-Whitney U test
      req(input$ttest_group)
      mw_test <- wilcox.test(df[[input$ttest_var]] ~ df[[input$ttest_group]])
      print(mw_test)
    }
  })
  
  output$download_ttest <- downloadHandler(
    filename = function() {
      "t_test_results.docx"
    },
    content = function(file) {
      df <- data()
      
      if(input$ttest_type == "one.sample") {
        t_test <- t.test(df[[input$ttest_var]], mu = input$ttest_mu)
        results <- data.frame(
          Test = "One Sample t-test",
          Statistic = t_test$statistic,
          DF = t_test$parameter,
          PValue = t_test$p.value,
          Estimate = t_test$estimate,
          ConfLow = t_test$conf.int[1],
          ConfHigh = t_test$conf.int[2]
        )
      } else if(input$ttest_type == "two.sample") {
        t_test <- t.test(df[[input$ttest_var]] ~ df[[input$ttest_group]], 
                         var.equal = input$ttest_var_equal)
        results <- data.frame(
          Test = "Two Sample t-test",
          Statistic = t_test$statistic,
          DF = t_test$parameter,
          PValue = t_test$p.value,
          Mean1 = t_test$estimate[1],
          Mean2 = t_test$estimate[2],
          ConfLow = t_test$conf.int[1],
          ConfHigh = t_test$conf.int[2]
        )
      } else if(input$ttest_type == "paired") {
        t_test <- t.test(df[[input$ttest_var]], df[[input$ttest_group]], paired = TRUE)
        results <- data.frame(
          Test = "Paired t-test",
          Statistic = t_test$statistic,
          DF = t_test$parameter,
          PValue = t_test$p.value,
          MeanDiff = t_test$estimate,
          ConfLow = t_test$conf.int[1],
          ConfHigh = t_test$conf.int[2]
        )
      } else {
        mw_test <- wilcox.test(df[[input$ttest_var]] ~ df[[input$ttest_group]])
        results <- data.frame(
          Test = "Mann-Whitney U Test",
          Statistic = mw_test$statistic,
          PValue = mw_test$p.value
        )
      }
      
      ft <- flextable::flextable(results)
      doc <- read_docx() %>% 
        officer::body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # ANOVA/Kruskal-Wallis
  output$anova_results <- renderPrint({
    req(input$anova_var, input$anova_group)
    df <- data()
    
    if(input$anova_type == "anova") {
      # One-way ANOVA
      anova_model <- aov(df[[input$anova_var]] ~ as.factor(df[[input$anova_group]]))
      cat("ANOVA Results:\n")
      print(summary(anova_model))
    } else {
      # Kruskal-Wallis test
      kw_test <- kruskal.test(df[[input$anova_var]] ~ as.factor(df[[input$anova_group]]))
      cat("Kruskal-Wallis Test Results:\n")
      print(kw_test)
    }
  })
  
  output$posthoc_results <- renderPrint({
    req(input$anova_var, input$anova_group, input$anova_posthoc)
    df <- data()
    
    if(input$anova_posthoc_type == "tukey") {
      # Tukey HSD post-hoc test
      anova_model <- aov(df[[input$anova_var]] ~ as.factor(df[[input$anova_group]]))
      cat("\nTukey HSD Post-hoc Test:\n")
      print(TukeyHSD(anova_model))
    } else {
      # Dunn's test
      cat("\nDunn's Test (Post-hoc for Kruskal-Wallis):\n")
      dunn_result <- dunn.test::dunn.test(df[[input$anova_var]], 
                                          as.factor(df[[input$anova_group]]),
                                          method = "bonferroni")
      print(dunn_result)
    }
  })
  
  output$download_anova <- downloadHandler(
    filename = function() {
      "anova_results.docx"
    },
    content = function(file) {
      df <- data()
      
      if(input$anova_type == "anova") {
        anova_model <- aov(df[[input$anova_var]] ~ as.factor(df[[input$anova_group]]))
        results <- data.frame(
          Test = "One-way ANOVA",
          DF = summary(anova_model)[[1]]["Df"],
          FValue = summary(anova_model)[[1]]["F value"],
          PValue = summary(anova_model)[[1]]["Pr(>F)"]
        )
      } else {
        kw_test <- kruskal.test(df[[input$anova_var]] ~ as.factor(df[[input$anova_group]]))
        results <- data.frame(
          Test = "Kruskal-Wallis Test",
          ChiSquared = kw_test$statistic,
          DF = kw_test$parameter,
          PValue = kw_test$p.value
        )
      }
      
      ft <- flextable::flextable(results)
      doc <- read_docx() %>% 
        officer::body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Chi-square test
  output$chisq_results <- renderPrint({
    req(input$chisq_var1, input$chisq_var2)
    df <- data()
    
    # Create contingency table
    tab <- table(df[[input$chisq_var1]], df[[input$chisq_var2]])
    
    # Perform chi-square test
    chisq_test <- chisq.test(tab)
    
    cat("Chi-square Test Results:\n")
    print(chisq_test)
    
    if(input$chisq_expected) {
      cat("\nExpected Counts:\n")
      print(chisq_test$expected)
    }
    
    if(input$chisq_residuals) {
      cat("\nPearson Residuals:\n")
      print(chisq_test$residuals)
    }
  })
  
  output$chisq_table <- renderDT({
    req(input$chisq_var1, input$chisq_var2)
    df <- data()
    
    # Create and display contingency table
    tab <- as.data.frame.matrix(table(df[[input$chisq_var1]], df[[input$chisq_var2]]))
    datatable(tab)
  })
  
  output$download_chisq <- downloadHandler(
    filename = function() {
      "chi_square_test.docx"
    },
    content = function(file) {
      df <- data()
      
      tab <- table(df[[input$chisq_var1]], df[[input$chisq_var2]])
      chisq_test <- chisq.test(tab)
      
      # Create results table
      results <- data.frame(
        Test = "Chi-square Test",
        Statistic = chisq_test$statistic,
        DF = chisq_test$parameter,
        PValue = chisq_test$p.value
      )
      
      ft <- flextable::flextable(results)
      doc <- read_docx() %>% 
        officer::body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Navigation buttons
  observeEvent(input$goto_data, {
    updateNavbarPage(session, "nav", selected = "Data Upload")
  })
  
  observeEvent(input$goto_descriptive, {
    updateNavbarPage(session, "nav", selected = "Descriptive Analysis")
  })
  
  observeEvent(input$goto_disease, {
    updateNavbarPage(session, "nav", selected = "Disease Frequency")
  })
  
  observeEvent(input$goto_statistics, {
    updateNavbarPage(session, "nav", selected = "Statistical Associations")
  })
}

# Run the application
shinyApp(ui = ui, server = server)
