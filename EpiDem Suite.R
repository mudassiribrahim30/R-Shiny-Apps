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
library(openxlsx)
library(janitor)
library(lubridate)
library(shinycssloaders)
library(shinyjs)
library(colourpicker)
library(flextable)
library(polycor) 
library(psych)   

# Custom CSS for better layout and animations
custom_css <- "
/* Animated app name in navbar */
.navbar-brand {
  animation: gentleMove 3s ease-in-out infinite;
  transform-origin: center;
}

@keyframes gentleMove {
  0% { transform: translateY(0) rotate(0deg); }
  25% { transform: translateY(-2px) rotate(0.5deg); }
  50% { transform: translateY(0) rotate(0deg); }
  75% { transform: translateY(-1px) rotate(-0.5deg); }
  100% { transform: translateY(0) rotate(0deg); }
}

/* Professional font settings */
body {
  font-family: 'Segoe UI', 'Roboto', 'Helvetica Neue', Arial, sans-serif;
  font-size: 15px;
  line-height: 1.6;
  color: #333333;
}

h1, h2, h3, h4, h5, h6 {
  font-weight: 600;
  color: #2c3e50;
}

.main-header .navbar-brand {
  font-weight: 700;
  font-size: 1.4rem;
  letter-spacing: 0.5px;
}

/* Improved sidebar and content panels */
.sidebar-panel {
  background: #f8f9fa;
  border-right: 1px solid #e9ecef;
  padding: 25px 20px;
  height: calc(100vh - 80px);
  overflow-y: auto;
  position: sticky;
  top: 0;
  box-shadow: 2px 0 10px rgba(0,0,0,0.05);
}

.main-panel {
  padding: 25px;
  height: calc(100vh - 80px);
  overflow-y: auto;
  background: white;
}

/* Enhanced card design */
.card {
  margin-bottom: 1.5rem;
  border: 1px solid #e9ecef;
  border-radius: 8px;
  box-shadow: 0 2px 4px rgba(0,0,0,0.05);
  transition: all 0.3s ease;
  background: white;
}

.card:hover {
  transform: translateY(-3px);
  box-shadow: 0 6px 15px rgba(0,0,0,0.1);
}

.card-header {
  background-color: #f8f9fa;
  border-bottom: 1px solid #e9ecef;
  padding: 1rem 1.25rem;
  font-weight: 600;
}

.card-body {
  padding: 1.5rem;
}

/* Button improvements */
.btn-action {
  font-weight: 500;
  padding: 0.5rem 1.25rem;
  border-radius: 6px;
  transition: all 0.3s ease;
  border: none;
  box-shadow: 0 2px 4px rgba(0,0,0,0.1);
}

.btn-action:hover {
  transform: translateY(-2px);
  box-shadow: 0 4px 8px rgba(0,0,0,0.15);
}

.btn-action:active {
  transform: translateY(0);
  box-shadow: 0 1px 2px rgba(0,0,0,0.1);
}

/* Form control improvements */
.form-control, .selectize-input {
  border-radius: 6px;
  border: 1px solid #ced4da;
  padding: 0.5rem 0.75rem;
  font-size: 0.95rem;
}

.form-control:focus, .selectize-input.focus {
  border-color: #0C5EA8;
  box-shadow: 0 0 0 0.2rem rgba(12, 94, 168, 0.25);
}

/* Better selectize dropdown */
.selectize-dropdown {
  border-radius: 6px;
  border: 1px solid #ced4da;
  box-shadow: 0 4px 12px rgba(0,0,0,0.1);
}

/* Improved tabs */
.nav-tabs {
  border-bottom: 2px solid #dee2e6;
  margin-bottom: 1.5rem;
}

.nav-tabs .nav-link {
  border: none;
  border-radius: 6px 6px 0 0;
  padding: 0.75rem 1.5rem;
  font-weight: 500;
  color: #6c757d;
  margin-right: 2px;
  transition: all 0.2s;
}

.nav-tabs .nav-link:hover {
  border-color: transparent;
  color: #0C5EA8;
  background-color: rgba(12, 94, 168, 0.05);
}

.nav-tabs .nav-link.active {
  color: #0C5EA8;
  background-color: rgba(12, 94, 168, 0.1);
  border-bottom: 3px solid #0C5EA8;
  font-weight: 600;
}

/* Table improvements */
.dataTables_wrapper {
  font-size: 0.9rem;
}

.dataTables_wrapper .dataTables_length,
.dataTables_wrapper .dataTables_filter {
  margin-bottom: 1rem;
}

/* Alert improvements */
.alert {
  border-radius: 8px;
  border: none;
  box-shadow: 0 2px 4px rgba(0,0,0,0.05);
  padding: 1rem 1.25rem;
}

.alert-success {
  background-color: #d4edda;
  color: #155724;
  border-left: 4px solid #28a745;
}

.alert-danger {
  background-color: #f8d7da;
  color: #721c24;
  border-left: 4px solid #dc3545;
}

.alert-warning {
  background-color: #fff3cd;
  color: #856404;
  border-left: 4px solid #ffc107;
}

.alert-info {
  background-color: #d1ecf1;
  color: #0c5460;
  border-left: 4px solid #17a2b8;
}

/* Loading spinner improvements */
.loading-container {
  display: flex;
  justify-content: center;
  align-items: center;
  padding: 3rem;
}

.loading-spinner {
  border: 4px solid #f3f3f3;
  border-top: 4px solid #0C5EA8;
  border-radius: 50%;
  width: 50px;
  height: 50px;
  animation: spin 1.5s linear infinite;
}

@keyframes spin {
  0% { transform: rotate(0deg); }
  100% { transform: rotate(360deg); }
}

/* Section headers */
.section-header {
  margin-bottom: 1.5rem;
  padding-bottom: 0.75rem;
  border-bottom: 2px solid #e9ecef;
  color: #2c3e50;
  font-weight: 600;
}

/* Input group improvements */
.input-group-prepend .input-group-text,
.input-group-append .input-group-text {
  background-color: #f8f9fa;
  border-color: #ced4da;
  font-weight: 500;
}

/* Better plot containers */
.plot-container {
  background: white;
  border-radius: 8px;
  padding: 1.5rem;
  border: 1px solid #e9ecef;
  box-shadow: 0 2px 4px rgba(0,0,0,0.05);
  margin-bottom: 1.5rem;
}

/* Tooltip improvements */
.tooltip {
  font-size: 0.85rem;
  max-width: 300px;
}

/* Scrollbar styling */
::-webkit-scrollbar {
  width: 10px;
}

::-webkit-scrollbar-track {
  background: #f1f1f1;
  border-radius: 10px;
}

::-webkit-scrollbar-thumb {
  background: #0C5EA8;
  border-radius: 10px;
}

::-webkit-scrollbar-thumb:hover {
  background: #094580;
}

/* Correlation analysis specific styles */
.correlation-plot {
  border: 2px solid #e9ecef;
  border-radius: 8px;
  padding: 15px;
  background: white;
  margin-top: 20px;
}

.correlation-results {
  font-family: 'Courier New', monospace;
  background: #f8f9fa;
  padding: 15px;
  border-radius: 6px;
  margin-top: 15px;
}

.coefficient-highlight {
  font-weight: bold;
  color: #0C5EA8;
  font-size: 1.1em;
}

.scatter-point {
  transition: all 0.3s ease;
}

.scatter-point:hover {
  stroke-width: 3;
  stroke: #000;
}

/* Dashboard value boxes */
.value-box {
  background: linear-gradient(135deg, #0C5EA8 0%, #094580 100%);
  color: white;
  border-radius: 8px;
  padding: 20px;
  margin-bottom: 20px;
  box-shadow: 0 4px 6px rgba(0,0,0,0.1);
}

.value-box .value {
  font-size: 2.5rem;
  font-weight: 700;
  line-height: 1;
}

.value-box .caption {
  font-size: 1rem;
  opacity: 0.9;
  margin-top: 5px;
}

/* Home page specific styles */
.jumbotron-home {
  background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
  padding: 3rem 2rem;
  margin-bottom: 2rem;
  border-radius: 12px;
  border: none;
  box-shadow: 0 4px 15px rgba(0,0,0,0.05);
}

.feature-card {
  height: 100%;
  border: 1px solid #e9ecef;
  border-radius: 10px;
  transition: all 0.3s ease;
  background: white;
}

.feature-card:hover {
  transform: translateY(-5px);
  box-shadow: 0 8px 25px rgba(0,0,0,0.15);
  border-color: #0C5EA8;
}

.feature-icon {
  font-size: 2.5rem;
  color: #0C5EA8;
  margin-bottom: 1rem;
}

/* Help text styling */
.help-block {
  font-size: 0.85rem;
  color: #6c757d;
  margin-top: 0.25rem;
}

/* Required field indicators */
.required:after {
  content: ' *';
  color: #dc3545;
}

/* Better spacing for form groups */
.form-group {
  margin-bottom: 1.25rem;
}

/* Checkbox and radio improvements */
.checkbox, .radio {
  margin-top: 0.5rem;
  margin-bottom: 0.5rem;
}

.checkbox label, .radio label {
  font-weight: 500;
}

/* Slider improvements */
.irs--shiny .irs-bar {
  background: #0C5EA8;
}

.irs--shiny .irs-handle {
  border: 3px solid #0C5EA8;
}

/* Footer improvements */
.footer {
  text-align: center;
  padding: 15px;
  background-color: #f8f9fa;
  border-top: 1px solid #dee2e6;
  color: #6c757d;
  font-size: 0.9rem;
}

/* Responsive adjustments */
@media (max-width: 768px) {
  .sidebar-panel, .main-panel {
    height: auto;
    position: static;
  }
  
  .sidebar-panel {
    border-right: none;
    border-bottom: 1px solid #dee2e6;
  }
}
"

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

# Helper function to handle labelled data
convert_labelled <- function(df) {
  df <- df %>%
    mutate(across(where(haven::is.labelled), as_factor)) %>%
    mutate(across(where(is.factor), as.character)) %>%
    mutate(across(where(is.character), ~ifelse(. == "", NA, .))) %>%
    mutate(across(everything(), ~ifelse(is.na(.), NA, .)))
  return(df)
}

# UI
ui <- navbarPage(
  title = div(
    span("📈 ", style = "font-size: 1.2em; color: white;"),
    span("EpiDem Suite™", 
         style = "font-weight: 700; letter-spacing: 0.5px; 
                  color: white; /* Changed from gradient to solid white */
                  text-shadow: 1px 1px 2px rgba(0,0,0,0.5); /* Added shadow for better visibility */
                  ")
  ),
  footer = div(
    class = "footer",
    "Copyright © Mudasir Mohammed Ibrahim (mudassiribrahim30@gmail.com)"
  ),
  windowTitle = "EpiDem Suite - Epidemiological Analysis Tool",
  id = "nav",
  header = tags$head(
    tags$style(HTML(custom_css)),
    useShinyjs(),
    tags$script(HTML("
      $(document).ready(function() {
        // Add animation to navbar brand
        $('.navbar-brand').css({
          'animation': 'gentleMove 3s ease-in-out infinite',
          'transform-origin': 'center'
        });
        
        // Smooth scrolling for anchor links
        $('a[href^=\"#\"]').on('click', function(e) {
          if(this.hash !== '') {
            e.preventDefault();
            const hash = this.hash;
            $('html, body').animate({
              scrollTop: $(hash).offset().top - 70
            }, 800);
          }
        });
      });
    "))
  ),
  theme = bslib::bs_theme(
    version = 5, 
    bootswatch = "flatly",
    primary = "#0C5EA8",   # CDC Blue
    secondary = "#CD2026", # CDC Red
    success = "#5CB85C",
    base_font = bslib::font_google("Roboto"),
    heading_font = bslib::font_google("Roboto"),
    font_scale = 1.05
  ),
  
  # Home/About Tab
  tabPanel(
    "Home", 
    icon = icon("house", class = "fa-lg"),
    div(
      class = "container-fluid",
      div(
        class = "jumbotron jumbotron-home",
        h1("EpiDem Suite", class = "display-4", style = "font-weight: 700; color: #2c3e50;"),
        p(
          "A comprehensive epidemiological analysis platform for public health professionals.", 
          class = "lead",
          style = "color: #5a6268; font-size: 1.25rem;"
        ),
        hr(class = "my-4", style = "border-color: #dee2e6;"),
        p(
          "Upload your dataset and utilize the tools in the navigation bar to perform descriptive analyses, 
          calculate disease frequency measures, assess statistical associations, and build regression models.",
          style = "font-size: 1.05rem; line-height: 1.6;"
        ),
        p(
          "Supported file formats: CSV, Excel (.xls, .xlsx), Stata (.dta), SPSS (.sav), SAS (.sas7bdat), RData (.RData, .rda)",
          style = "font-size: 0.95rem; color: #6c757d;"
        ),
        p(
          "This tool follows established guidelines for epidemiological analysis. Developed by Mudasir Mohammed Ibrahim (BSc, RN)",
          style = "font-size: 0.9rem; color: #6c757d; font-style: italic;"
        ),
        actionButton(
          "goto_data", 
          "Get Started →", 
          class = "btn-primary btn-lg btn-action", 
          icon = icon("rocket"),
          style = "margin-top: 1rem; padding: 0.75rem 2rem; font-weight: 600;"
        )
      ),
      
      fluidRow(
        column(
          4,
          div(
            class = "card feature-card",
            div(
              class = "card-body text-center",
              div(class = "feature-icon", "📊"),
              h3("Descriptive Analysis", class = "card-title", style = "font-weight: 600;"),
              p(
                "Explore frequency distributions, summary statistics, and create epidemic curves to understand your data's basic characteristics.",
                class = "card-text",
                style = "color: #5a6268;"
              ),
              tags$ul(
                class = "text-left",
                style = "padding-left: 1.5rem;",
                tags$li("Frequency distributions and bar plots"),
                tags$li("Central tendency measures and histograms"),
                tags$li("Epidemic curves with customizable intervals"),
                tags$li("Spatial analysis and demographic breakdowns")
              ),
              actionButton(
                "goto_descriptive", 
                "Explore Descriptive Tools", 
                class = "btn-outline-primary btn-action",
                style = "margin-top: 1rem;"
              )
            )
          )
        ),
        column(
          4,
          div(
            class = "card feature-card",
            div(
              class = "card-body text-center",
              div(class = "feature-icon", "🦠"),
              h3("Disease Metrics", class = "card-title", style = "font-weight: 600;"),
              p(
                "Calculate incidence rates, prevalence, attack rates, and mortality measures to quantify disease burden in populations.",
                class = "card-text",
                style = "color: #5a6268;"
              ),
              tags$ul(
                class = "text-left",
                style = "padding-left: 1.5rem;",
                tags$li("Incidence rates and cumulative incidence"),
                tags$li("Prevalence calculations"),
                tags$li("Attack rates with confidence intervals"),
                tags$li("Mortality rates and case fatality rates"),
                tags$li("Life table analysis")
              ),
              actionButton(
                "goto_disease", 
                "Calculate Disease Metrics", 
                class = "btn-outline-primary btn-action",
                style = "margin-top: 1rem;"
              )
            )
          )
        ),
        column(
          4,
          div(
            class = "card feature-card",
            div(
              class = "card-body text-center",
              div(class = "feature-icon", "📈"),
              h3("Statistical Tools", class = "card-title", style = "font-weight: 600;"),
              p(
                "Perform association tests, regression analyses, and survival analysis to identify risk factors and measure effects.",
                class = "card-text",
                style = "color: #5a6268;"
              ),
              tags$ul(
                class = "text-left",
                style = "padding-left: 1.5rem;",
                tags$li("Risk ratios, odds ratios, and attributable risk"),
                tags$li("Poisson regression for count data"),
                tags$li("Survival analysis with Kaplan-Meier curves"),
                tags$li("Cox proportional hazards models"),
                tags$li("Linear and logistic regression"),
                tags$li("Chi-square tests with cross-tabulations"),
                tags$li("Correlation analysis (Pearson, Kendall, Spearman)")
              ),
              actionButton(
                "goto_statistics", 
                "Use Statistical Tools", 
                class = "btn-outline-primary btn-action",
                style = "margin-top: 1rem;"
              )
            )
          )
        )
      )
    )
  ),
  
  # Data Upload Tab
  tabPanel(
    "Data Upload", 
    icon = icon("database", class = "fa-lg"),
    fluidRow(
      column(
        4, 
        class = "sidebar-panel",
        div(
          class = "section-header",
          h4("Data Upload Settings", style = "margin: 0;")
        ),
        fileInput(
          "file1", 
          "Choose File",
          accept = c(".csv", ".xls", ".xlsx", ".sav", ".dta", ".sas7bdat", ".RData", ".rda"),
          multiple = FALSE,
          width = "100%",
          buttonLabel = "Browse...",
          placeholder = "No file selected"
        ),
        helpText(
          "Supported formats: CSV, Excel, Stata, SPSS, SAS, RData",
          class = "help-block"
        ),
        tags$hr(),
        conditionalPanel(
          condition = "input.file1 && input.file1.type.includes('csv')",
          checkboxInput("header", "Header", TRUE, width = "100%"),
          selectInput(
            "sep", 
            "Separator",
            choices = c(Comma = ",", Semicolon = ";", Tab = "\t", Space = " "),
            selected = ","
          ),
          selectInput(
            "quote", 
            "Quote",
            choices = c(None = "", "Double Quote" = '"', "Single Quote" = "'"),
            selected = '"'
          )
        ),
        conditionalPanel(
          condition = "input.file1 && (input.file1.type.includes('xls') || input.file1.type.includes('xlsx'))",
          numericInput("sheet", "Sheet Number", value = 1, min = 1, width = "100%")
        ),
        actionButton(
          "load_data", 
          "Load Data", 
          class = "btn-primary btn-action",
          icon = icon("upload"),
          width = "100%",
          style = "margin-top: 10px;"
        ),
        tags$hr(),
        uiOutput("missing_values_alert")
      ),
      column(
        8, 
        class = "main-panel",
        div(
          class = "section-header",
          h4("Data Preview", style = "margin: 0;")
        ),
        withSpinner(
          DTOutput("contents"), 
          type = 4, 
          color = "#0C5EA8",
          size = 1.5
        ),
        div(
          class = "section-header",
          style = "margin-top: 30px;",
          h4("Data Summary", style = "margin: 0;")
        ),
        withSpinner(
          verbatimTextOutput("data_summary"), 
          type = 4, 
          color = "#0C5EA8",
          size = 1.5
        )
      )
    )
  ),
  
  # Descriptive Analysis Tab
  tabPanel(
    "Descriptive Analysis", 
    icon = icon("chart-bar", class = "fa-lg"),
    tabsetPanel(
      id = "desc_tabs",
      tabPanel(
        "Frequency Analysis",
        fluidRow(
          column(
            4, 
            class = "sidebar-panel",
            div(
              class = "section-header",
              h4("Frequency Analysis Settings", style = "margin: 0;")
            ),
            selectInput(
              "freq_var", 
              "Select Variable:", 
              choices = NULL,
              width = "100%"
            ),
            colourpicker::colourInput(
              "freq_color", 
              "Bar Color:", 
              value = "#0C5EA8",
              showColour = "background",
              allowTransparent = TRUE
            ),
            textInput(
              "freq_title", 
              "Plot Title:", 
              value = "Frequency Distribution",
              width = "100%"
            ),
            fluidRow(
              column(
                6,
                numericInput(
                  "freq_title_size", 
                  "Title Size:", 
                  value = 16, 
                  min = 8, 
                  max = 24,
                  width = "100%"
                )
              ),
              column(
                6,
                numericInput(
                  "freq_x_size", 
                  "X-axis Label Size:", 
                  value = 14, 
                  min = 8, 
                  max = 20,
                  width = "100%"
                )
              )
            ),
            radioButtons(
              "freq_type", 
              "Download Format:",
              choices = c("Table", "Plot"), 
              selected = "Table",
              inline = TRUE
            ),
            actionButton(
              "run_freq", 
              "Run Analysis", 
              class = "btn-primary btn-action",
              icon = icon("play"),
              width = "100%",
              style = "margin-top: 10px;"
            ),
            downloadButton(
              "download_freq", 
              "Download Results",
              class = "btn-success btn-action",
              width = "100%",
              style = "margin-top: 10px;"
            )
          ),
          column(
            8, 
            class = "main-panel",
            withSpinner(
              DTOutput("freq_table"), 
              type = 4, 
              color = "#0C5EA8",
              size = 1.5
            ),
            div(
              class = "plot-container",
              withSpinner(
                plotlyOutput("freq_plot"), 
                type = 4, 
                color = "#0C5EA8",
                size = 1.5
              )
            )
          )
        )
      ),
      tabPanel(
        "Central Tendency",
        fluidRow(
          column(
            4, 
            class = "sidebar-panel",
            div(
              class = "section-header",
              h4("Central Tendency Settings", style = "margin: 0;")
            ),
            selectInput(
              "central_var", 
              "Select Numeric Variable:", 
              choices = NULL,
              width = "100%"
            ),
            selectInput(
              "central_stratify", 
              "Stratify by (Optional):", 
              choices = c("None" = "none"),
              width = "100%"
            ),
            colourpicker::colourInput(
              "central_color", 
              "Plot Color:", 
              value = "#0C5EA8",
              showColour = "background",
              allowTransparent = TRUE
            ),
            textInput(
              "central_title", 
              "Plot Title:", 
              value = "Central Tendency",
              width = "100%"
            ),
            numericInput(
              "central_bins", 
              "Number of Bins:", 
              value = 30, 
              min = 5, 
              max = 100,
              width = "100%"
            ),
            fluidRow(
              column(
                4,
                numericInput(
                  "central_title_size", 
                  "Title Size:", 
                  value = 16, 
                  min = 8, 
                  max = 24,
                  width = "100%"
                )
              ),
              column(
                4,
                numericInput(
                  "central_x_size", 
                  "X-axis Label Size:", 
                  value = 14, 
                  min = 8, 
                  max = 20,
                  width = "100%"
                )
              ),
              column(
                4,
                numericInput(
                  "central_y_size", 
                  "Y-axis Label Size:", 
                  value = 14, 
                  min = 8, 
                  max = 20,
                  width = "100%"
                )
              )
            ),
            radioButtons(
              "central_type", 
              "Download Format:",
              choices = c("Table", "Boxplot", "Histogram"), 
              selected = "Table",
              inline = TRUE
            ),
            actionButton(
              "run_central", 
              "Run Analysis", 
              class = "btn-primary btn-action",
              icon = icon("play"),
              width = "100%",
              style = "margin-top: 10px;"
            ),
            downloadButton(
              "download_central", 
              "Download Results",
              class = "btn-success btn-action",
              width = "100%",
              style = "margin-top: 10px;"
            )
          ),
          column(
            8, 
            class = "main-panel",
            withSpinner(
              DTOutput("central_table"), 
              type = 4, 
              color = "#0C5EA8",
              size = 1.5
            ),
            div(
              class = "plot-container",
              withSpinner(
                plotlyOutput("central_boxplot"), 
                type = 4, 
                color = "#0C5EA8",
                size = 1.5
              )
            ),
            div(
              class = "plot-container",
              withSpinner(
                plotlyOutput("central_histogram"), 
                type = 4, 
                color = "#0C5EA8",
                size = 1.5
              )
            )
          )
        )
      ),
      tabPanel(
        "Epidemic Curve",
        fluidRow(
          column(
            4, 
            class = "sidebar-panel",
            div(
              class = "section-header",
              h4("Epidemic Curve Settings", style = "margin: 0;")
            ),
            selectInput(
              "epi_date", 
              "Date Variable:", 
              choices = NULL,
              width = "100%"
            ),
            selectInput(
              "epi_case", 
              "Case Count Variable:", 
              choices = NULL,
              width = "100%"
            ),
            selectInput(
              "epi_group", 
              "Grouping Variable (Optional):", 
              choices = c("None" = "none"),
              width = "100%"
            ),
            selectInput(
              "epi_interval", 
              "Time Interval:",
              choices = c("Day", "Week", "Month", "Year"), 
              selected = "Week",
              width = "100%"
            ),
            colourpicker::colourInput(
              "epi_color", 
              "Bar Color:", 
              value = "#CD2026",
              showColour = "background",
              allowTransparent = TRUE
            ),
            textInput(
              "epi_title", 
              "Plot Title:", 
              value = "Epidemic Curve",
              width = "100%"
            ),
            fluidRow(
              column(
                4,
                numericInput(
                  "epi_title_size", 
                  "Title Size:", 
                  value = 16, 
                  min = 8, 
                  max = 24,
                  width = "100%"
                )
              ),
              column(
                4,
                numericInput(
                  "epi_x_size", 
                  "X-axis Label Size:", 
                  value = 14, 
                  min = 8, 
                  max = 20,
                  width = "100%"
                )
              ),
              column(
                4,
                numericInput(
                  "epi_y_size", 
                  "Y-axis Label Size:", 
                  value = 14, 
                  min = 8, 
                  max = 20,
                  width = "100%"
                )
              )
            ),
            actionButton(
              "run_epi", 
              "Run Analysis", 
              class = "btn-primary btn-action",
              icon = icon("play"),
              width = "100%",
              style = "margin-top: 10px;"
            ),
            downloadButton(
              "download_epicurve", 
              "Download Plot",
              class = "btn-success btn-action",
              width = "100%",
              style = "margin-top: 10px;"
            ),
            downloadButton(
              "download_epi_data", 
              "Download Data",
              class = "btn-info btn-action",
              width = "100%",
              style = "margin-top: 10px;"
            )
          ),
          column(
            8, 
            class = "main-panel",
            withSpinner(
              uiOutput("epi_summary"), 
              type = 4, 
              color = "#0C5EA8",
              size = 1.5
            ),
            div(
              class = "plot-container",
              withSpinner(
                plotlyOutput("epi_curve"), 
                type = 4, 
                color = "#0C5EA8",
                size = 1.5
              )
            )
          )
        )
      )
    )
  ),
  
  # Disease Frequency Tab
  tabPanel(
    "Disease Frequency", 
    icon = icon("viruses", class = "fa-lg"),
    tabsetPanel(
      id = "disease_tabs",
      tabPanel(
        "Incidence/Prevalence",
        fluidRow(
          column(
            4, 
            class = "sidebar-panel",
            div(
              class = "section-header",
              h4("Incidence/Prevalence Settings", style = "margin: 0;")
            ),
            selectInput(
              "ip_measure", 
              "Select Measure:",
              choices = c("Incidence Rate", "Cumulative Incidence", "Prevalence"),
              selected = "Incidence Rate",
              width = "100%"
            ),
            selectInput(
              "ip_case", 
              "Case Count Variable:", 
              choices = NULL,
              width = "100%"
            ),
            selectInput(
              "ip_pop", 
              "Population Variable:", 
              choices = NULL,
              width = "100%"
            ),
            conditionalPanel(
              condition = "input.ip_measure == 'Incidence Rate'",
              selectInput(
                "ip_time", 
                "Person-Time Variable:", 
                choices = NULL,
                width = "100%"
              )
            ),
            actionButton(
              "run_ip", 
              "Calculate", 
              class = "btn-primary btn-action",
              icon = icon("calculator"),
              width = "100%",
              style = "margin-top: 10px;"
            ),
            downloadButton(
              "download_ip", 
              "Download Results",
              class = "btn-success btn-action",
              width = "100%",
              style = "margin-top: 10px;"
            )
          ),
          column(
            8, 
            class = "main-panel",
            withSpinner(
              DTOutput("ip_table"), 
              type = 4, 
              color = "#0C5EA8",
              size = 1.5
            )
          )
        )
      )
    )
  ),
  
  # Statistical Associations Tab
  tabPanel(
    "Statistical Associations", 
    icon = icon("calculator", class = "fa-lg"),
    tabsetPanel(
      id = "assoc_tabs",
      tabPanel(
        "Risk Ratios",
        fluidRow(
          column(
            4, 
            class = "sidebar-panel",
            div(
              class = "section-header",
              h4("Risk Ratio Settings", style = "margin: 0;")
            ),
            selectInput(
              "rr_outcome", 
              "Outcome Variable:", 
              choices = NULL,
              width = "100%"
            ),
            selectInput(
              "rr_exposure", 
              "Exposure Variable:", 
              choices = NULL,
              width = "100%"
            ),
            uiOutput("rr_outcome_levels_ui"),
            uiOutput("rr_exposure_levels_ui"),
            sliderInput(
              "rr_conf", 
              "Confidence Level:", 
              min = 0.90, 
              max = 0.99, 
              value = 0.95, 
              step = 0.01,
              width = "100%"
            ),
            actionButton(
              "run_rr", 
              "Calculate", 
              class = "btn-primary btn-action",
              icon = icon("calculator"),
              width = "100%",
              style = "margin-top: 10px;"
            ),
            downloadButton(
              "download_rr", 
              "Download Results",
              class = "btn-success btn-action",
              width = "100%",
              style = "margin-top: 10px;"
            )
          ),
          column(
            8, 
            class = "main-panel",
            withSpinner(
              verbatimTextOutput("rr_results"), 
              type = 4, 
              color = "#0C5EA8",
              size = 1.5
            ),
            withSpinner(
              DTOutput("rr_table"), 
              type = 4, 
              color = "#0C5EA8",
              size = 1.5
            ),
            withSpinner(
              verbatimTextOutput("rr_interpretation"), 
              type = 4, 
              color = "#0C5EA8",
              size = 1.5
            )
          )
        )
      ),
      tabPanel(
        "Odds Ratios",
        fluidRow(
          column(
            4, 
            class = "sidebar-panel",
            div(
              class = "section-header",
              h4("Odds Ratio Settings", style = "margin: 0;")
            ),
            selectInput(
              "or_outcome", 
              "Outcome Variable:", 
              choices = NULL,
              width = "100%"
            ),
            selectInput(
              "or_exposure", 
              "Exposure Variable:", 
              choices = NULL,
              width = "100%"
            ),
            sliderInput(
              "or_conf", 
              "Confidence Level:", 
              min = 0.90, 
              max = 0.99, 
              value = 0.95, 
              step = 0.01,
              width = "100%"
            ),
            actionButton(
              "run_or", 
              "Calculate", 
              class = "btn-primary btn-action",
              icon = icon("calculator"),
              width = "100%",
              style = "margin-top: 10px;"
            ),
            downloadButton(
              "download_or", 
              "Download Results",
              class = "btn-success btn-action",
              width = "100%",
              style = "margin-top: 10px;"
            )
          ),
          column(
            8, 
            class = "main-panel",
            withSpinner(
              verbatimTextOutput("or_results"), 
              type = 4, 
              color = "#0C5EA8",
              size = 1.5
            ),
            withSpinner(
              DTOutput("or_table"), 
              type = 4, 
              color = "#0C5EA8",
              size = 1.5
            )
          )
        )
      )
    )
  ),
  
  # Survival Analysis Tab
  tabPanel(
    "Survival Analysis", 
    icon = icon("heartbeat", class = "fa-lg"),
    tabsetPanel(
      id = "surv_tabs",
      tabPanel(
        "Kaplan-Meier",
        tabsetPanel(
          tabPanel(
            "Analysis Settings",
            fluidRow(
              column(
                4, 
                class = "sidebar-panel",
                div(
                  class = "section-header",
                  h4("Kaplan-Meier Settings", style = "margin: 0;")
                ),
                selectInput(
                  "surv_time", 
                  "Time Variable:", 
                  choices = NULL,
                  width = "100%"
                ),
                selectInput(
                  "surv_event", 
                  "Event Variable:", 
                  choices = NULL,
                  width = "100%"
                ),
                selectInput(
                  "surv_group", 
                  "Grouping Variable (Optional):", 
                  choices = c("None" = "none"),
                  width = "100%"
                ),
                actionButton(
                  "run_surv", 
                  "Run Analysis", 
                  class = "btn-primary btn-action",
                  icon = icon("play"),
                  width = "100%",
                  style = "margin-top: 10px;"
                ),
                downloadButton(
                  "download_surv_data", 
                  "Download Data",
                  class = "btn-info btn-action",
                  width = "100%",
                  style = "margin-top: 10px;"
                )
              ),
              column(
                8, 
                class = "main-panel",
                withSpinner(
                  uiOutput("surv_stats"), 
                  type = 4, 
                  color = "#0C5EA8",
                  size = 1.5
                ),
                withSpinner(
                  verbatimTextOutput("surv_summary"), 
                  type = 4, 
                  color = "#0C5EA8",
                  size = 1.5
                )
              )
            )
          ),
          tabPanel(
            "Plot Options",
            fluidRow(
              column(
                4, 
                class = "sidebar-panel",
                div(
                  class = "section-header",
                  h4("Display Options", style = "margin: 0;")
                ),
                checkboxInput(
                  "surv_ci", 
                  "Show Confidence Intervals", 
                  TRUE,
                  width = "100%"
                ),
                checkboxInput(
                  "surv_risktable", 
                  "Show Risk Table", 
                  TRUE,
                  width = "100%"
                ),
                checkboxInput(
                  "surv_median", 
                  "Show Median Survival Lines", 
                  TRUE,
                  width = "100%"
                ),
                checkboxInput(
                  "surv_labels", 
                  "Show Survival Probability Labels", 
                  TRUE,
                  width = "100%"
                ),
                br(),
                div(
                  class = "section-header",
                  h4("Color Settings", style = "margin: 0;")
                ),
                colourpicker::colourInput(
                  "surv_color1", 
                  "Group 1 Color:", 
                  value = "#0C5EA8",
                  showColour = "background"
                ),
                colourpicker::colourInput(
                  "surv_color2", 
                  "Group 2 Color:", 
                  value = "#CD2026",
                  showColour = "background"
                ),
                colourpicker::colourInput(
                  "surv_color3", 
                  "Group 3 Color:", 
                  value = "#5CB85C",
                  showColour = "background"
                ),
                br(),
                div(
                  class = "section-header",
                  h4("Text Settings", style = "margin: 0;")
                ),
                textInput(
                  "surv_title", 
                  "Plot Title:", 
                  value = "Kaplan-Meier Curve",
                  width = "100%"
                ),
                textInput(
                  "surv_xlab", 
                  "X-axis Label:", 
                  value = "Time",
                  width = "100%"
                ),
                textInput(
                  "surv_ylab", 
                  "Y-axis Label:", 
                  value = "Survival Probability",
                  width = "100%"
                ),
                fluidRow(
                  column(
                    6,
                    numericInput(
                      "surv_title_size", 
                      "Title Size:", 
                      value = 18, 
                      min = 8, 
                      max = 24,
                      width = "100%"
                    )
                  ),
                  column(
                    6,
                    numericInput(
                      "surv_label_size", 
                      "Label Size:", 
                      value = 5, 
                      min = 3, 
                      max = 10,
                      width = "100%"
                    )
                  )
                ),
                br(),
                downloadButton(
                  "download_surv", 
                  "Download Plot",
                  class = "btn-success btn-action",
                  width = "100%"
                )
              ),
              column(
                8, 
                class = "main-panel",
                div(
                  class = "plot-container",
                  withSpinner(
                    plotOutput("surv_plot", height = "600px"), 
                    type = 4, 
                    color = "#0C5EA8",
                    size = 1.5
                  )
                )
              )
            )
          ),
          tabPanel(
            "Life Table",
            fluidRow(
              column(
                4, 
                class = "sidebar-panel",
                div(
                  class = "section-header",
                  h4("Life Table Configuration", style = "margin: 0;")
                ),
                helpText(
                  "Specify time points (comma-separated) to evaluate survival probabilities:"
                ),
                textInput(
                  "life_table_times", 
                  "Time Points:", 
                  value = "0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0",
                  width = "100%"
                ),
                helpText(
                  "Example: 0.1, 0.25, 0.5, 0.75, 1.0 or specific time values like 30, 60, 90"
                ),
                numericInput(
                  "life_table_decimals", 
                  "Decimal Places:", 
                  value = 4, 
                  min = 2, 
                  max = 6,
                  width = "100%"
                ),
                actionButton(
                  "update_life_table", 
                  "Update Life Table", 
                  class = "btn-primary btn-action",
                  icon = icon("refresh"),
                  width = "100%",
                  style = "margin-top: 10px;"
                ),
                br(), br(),
                downloadButton(
                  "download_life_table", 
                  "Download Life Table (Excel)",
                  class = "btn-success btn-action",
                  width = "100%"
                )
              ),
              column(
                8, 
                class = "main-panel",
                h3("Life Table - Survival Probabilities at Specified Time Points"),
                helpText(
                  "This table shows the survival probabilities at the specified time points for each group."
                ),
                withSpinner(
                  DTOutput("life_table_output"), 
                  type = 4, 
                  color = "#0C5EA8",
                  size = 1.5
                ),
                br(),
                h4("Life Table Summary"),
                withSpinner(
                  verbatimTextOutput("life_table_summary"), 
                  type = 4, 
                  color = "#0C5EA8",
                  size = 1.5
                )
              )
            )
          )
        )
      ),
      tabPanel(
        "Cox Regression",
        fluidRow(
          column(
            4, 
            class = "sidebar-panel",
            div(
              class = "section-header",
              h4("Cox Regression Settings", style = "margin: 0;")
            ),
            selectInput(
              "cox_time", 
              "Time Variable:", 
              choices = NULL,
              width = "100%"
            ),
            selectInput(
              "cox_event", 
              "Event Variable:", 
              choices = NULL,
              width = "100%"
            ),
            selectizeInput(
              "cox_vars", 
              "Predictor Variables:", 
              choices = NULL, 
              multiple = TRUE,
              width = "100%",
              options = list(placeholder = 'Select variables...')
            ),
            uiOutput("cox_ref_ui"),
            sliderInput(
              "cox_conf", 
              "Confidence Level:", 
              min = 0.90, 
              max = 0.99, 
              value = 0.95, 
              step = 0.01,
              width = "100%"
            ),
            actionButton(
              "run_cox", 
              "Run Analysis", 
              class = "btn-primary btn-action",
              icon = icon("play"),
              width = "100%",
              style = "margin-top: 10px;"
            ),
            downloadButton(
              "download_cox", 
              "Download Results",
              class = "btn-success btn-action",
              width = "100%",
              style = "margin-top: 10px;"
            )
          ),
          column(
            8, 
            class = "main-panel",
            withSpinner(
              verbatimTextOutput("cox_summary"), 
              type = 4, 
              color = "#0C5EA8",
              size = 1.5
            ),
            withSpinner(
              DTOutput("cox_table"), 
              type = 4, 
              color = "#0C5EA8",
              size = 1.5
            )
          )
        )
      ),
      tabPanel(
        "Poisson Regression",
        fluidRow(
          column(
            4, 
            class = "sidebar-panel",
            div(
              class = "section-header",
              h4("Poisson Regression Settings", style = "margin: 0;")
            ),
            selectInput(
              "poisson_outcome", 
              "Outcome Variable (Count):", 
              choices = NULL,
              width = "100%"
            ),
            selectizeInput(
              "poisson_vars", 
              "Predictor Variables:", 
              choices = NULL, 
              multiple = TRUE,
              width = "100%",
              options = list(placeholder = 'Select variables...')
            ),
            uiOutput("poisson_ref_ui"),
            checkboxInput(
              "poisson_offset", 
              "Include Offset Variable", 
              FALSE,
              width = "100%"
            ),
            conditionalPanel(
              condition = "input.poisson_offset",
              selectInput(
                "poisson_offset_var", 
                "Offset Variable:", 
                choices = NULL,
                width = "100%"
              )
            ),
            sliderInput(
              "poisson_conf", 
              "Confidence Level:", 
              min = 0.90, 
              max = 0.99, 
              value = 0.95, 
              step = 0.01,
              width = "100%"
            ),
            actionButton(
              "run_poisson", 
              "Run Analysis", 
              class = "btn-primary btn-action",
              icon = icon("play"),
              width = "100%",
              style = "margin-top: 10px;"
            ),
            downloadButton(
              "download_poisson", 
              "Download Results",
              class = "btn-success btn-action",
              width = "100%",
              style = "margin-top: 10px;"
            )
          ),
          column(
            8, 
            class = "main-panel",
            withSpinner(
              verbatimTextOutput("poisson_summary"), 
              type = 4, 
              color = "#0C5EA8",
              size = 1.5
            ),
            withSpinner(
              DTOutput("poisson_table"), 
              type = 4, 
              color = "#0C5EA8",
              size = 1.5
            )
          )
        )
      )
    )
  ),
  
  # Regression Analysis Tab
  tabPanel(
    "Regression Analysis", 
    icon = icon("line-chart", class = "fa-lg"),
    tabsetPanel(
      id = "reg_tabs",
      tabPanel(
        "Linear Regression",
        fluidRow(
          column(
            4, 
            class = "sidebar-panel",
            div(
              class = "section-header",
              h4("Linear Regression Settings", style = "margin: 0;")
            ),
            selectInput(
              "linear_outcome", 
              "Outcome Variable:", 
              choices = NULL,
              width = "100%"
            ),
            selectizeInput(
              "linear_vars", 
              "Predictor Variables:", 
              choices = NULL, 
              multiple = TRUE,
              width = "100%",
              options = list(placeholder = 'Select variables...')
            ),
            uiOutput("linear_ref_ui"),
            checkboxInput(
              "linear_std", 
              "Show Standardized Coefficients", 
              FALSE,
              width = "100%"
            ),
            actionButton(
              "run_linear", 
              "Run Analysis", 
              class = "btn-primary btn-action",
              icon = icon("play"),
              width = "100%",
              style = "margin-top: 10px;"
            ),
            downloadButton(
              "download_linear", 
              "Download Results",
              class = "btn-success btn-action",
              width = "100%",
              style = "margin-top: 10px;"
            )
          ),
          column(
            8, 
            class = "main-panel",
            withSpinner(
              verbatimTextOutput("linear_summary"), 
              type = 4, 
              color = "#0C5EA8",
              size = 1.5
            ),
            withSpinner(
              DTOutput("linear_table"), 
              type = 4, 
              color = "#0C5EA8",
              size = 1.5
            )
          )
        )
      ),
      tabPanel(
        "Logistic Regression",
        fluidRow(
          column(
            4, 
            class = "sidebar-panel",
            div(
              class = "section-header",
              h4("Logistic Regression Settings", style = "margin: 0;")
            ),
            selectInput(
              "logistic_outcome", 
              "Outcome Variable:", 
              choices = NULL,
              width = "100%"
            ),
            uiOutput("logistic_outcome_ui"),
            selectizeInput(
              "logistic_vars", 
              "Predictor Variables:", 
              choices = NULL, 
              multiple = TRUE,
              width = "100%",
              options = list(placeholder = 'Select variables...')
            ),
            uiOutput("logistic_ref_ui"),
            checkboxInput(
              "logistic_or", 
              "Show Odds Ratios", 
              TRUE,
              width = "100%"
            ),
            actionButton(
              "run_logistic", 
              "Run Analysis", 
              class = "btn-primary btn-action",
              icon = icon("play"),
              width = "100%",
              style = "margin-top: 10px;"
            ),
            downloadButton(
              "download_logistic", 
              "Download Results",
              class = "btn-success btn-action",
              width = "100%",
              style = "margin-top: 10px;"
            )
          ),
          column(
            8, 
            class = "main-panel",
            withSpinner(
              verbatimTextOutput("logistic_summary"), 
              type = 4, 
              color = "#0C5EA8",
              size = 1.5
            ),
            withSpinner(
              DTOutput("logistic_table"), 
              type = 4, 
              color = "#0C5EA8",
              size = 1.5
            )
          )
        )
      )
    )
  ),
  
  # Statistical Tests Tab
  tabPanel(
    "Statistical Tests", 
    icon = icon("check-square", class = "fa-lg"),
    tabsetPanel(
      id = "test_tabs",
      tabPanel(
        "Chi-square Test",
        fluidRow(
          column(
            4, 
            class = "sidebar-panel",
            div(
              class = "section-header",
              h4("Chi-square Test Settings", style = "margin: 0;")
            ),
            selectInput(
              "chisq_var1", 
              "Variable 1:", 
              choices = NULL,
              width = "100%"
            ),
            selectInput(
              "chisq_var2", 
              "Variable 2:", 
              choices = NULL,
              width = "100%"
            ),
            checkboxInput(
              "chisq_expected", 
              "Show Expected Counts", 
              FALSE,
              width = "100%"
            ),
            checkboxInput(
              "chisq_residuals", 
              "Show Residuals", 
              FALSE,
              width = "100%"
            ),
            checkboxInput(
              "chisq_fisher", 
              "Include Fisher's Exact Test", 
              FALSE,
              width = "100%"
            ),
            checkboxInput(
              "chisq_rr", 
              "Include Relative Risk", 
              FALSE,
              width = "100%"
            ),
            actionButton(
              "run_chisq", 
              "Run Test", 
              class = "btn-primary btn-action",
              icon = icon("play"),
              width = "100%",
              style = "margin-top: 10px;"
            ),
            downloadButton(
              "download_chisq", 
              "Download Results",
              class = "btn-success btn-action",
              width = "100%",
              style = "margin-top: 10px;"
            )
          ),
          column(
            8, 
            class = "main-panel",
            withSpinner(
              verbatimTextOutput("chisq_results"), 
              type = 4, 
              color = "#0C5EA8",
              size = 1.5
            ),
            withSpinner(
              DTOutput("chisq_table"), 
              type = 4, 
              color = "#0C5EA8",
              size = 1.5
            ),
            div(
              class = "plot-container",
              withSpinner(
                plotOutput("chisq_plot"), 
                type = 4, 
                color = "#0C5EA8",
                size = 1.5
              )
            ),
            downloadButton(
              "download_chisq_plot", 
              "Download Plot",
              class = "btn-info btn-action",
              width = "100%",
              style = "margin-top: 10px;"
            )
          )
        )
      ),
      
      # Correlation Analysis Tab (NEW)
      tabPanel(
        "Correlation Analysis",
        fluidRow(
          column(
            4, 
            class = "sidebar-panel",
            div(
              class = "section-header",
              h4("Correlation Analysis Settings", style = "margin: 0;")
            ),
            selectInput(
              "cor_var1", 
              "Variable 1 (X-axis):", 
              choices = NULL,
              width = "100%"
            ),
            selectInput(
              "cor_var2", 
              "Variable 2 (Y-axis):", 
              choices = NULL,
              width = "100%"
            ),
            selectInput(
              "cor_method", 
              "Correlation Method:",
              choices = c(
                "Pearson" = "pearson",
                "Kendall" = "kendall", 
                "Spearman" = "spearman"
              ),
              selected = "pearson",
              width = "100%"
            ),
            checkboxInput(
              "cor_scatter", 
              "Show Scatter Plot", 
              TRUE,
              width = "100%"
            ),
            checkboxInput(
              "cor_fit", 
              "Add Trend Line", 
              TRUE,
              width = "100%"
            ),
            conditionalPanel(
              condition = "input.cor_scatter == true",
              colourpicker::colourInput(
                "cor_point_color", 
                "Point Color:", 
                value = "#0C5EA8",
                showColour = "background",
                allowTransparent = TRUE
              ),
              numericInput(
                "cor_point_size", 
                "Point Size:", 
                value = 3, 
                min = 1, 
                max = 10,
                width = "100%"
              ),
              numericInput(
                "cor_point_alpha", 
                "Point Transparency:", 
                value = 0.6, 
                min = 0.1, 
                max = 1,
                step = 0.1,
                width = "100%"
              )
            ),
            conditionalPanel(
              condition = "input.cor_fit == true && input.cor_scatter == true",
              colourpicker::colourInput(
                "cor_line_color", 
                "Trend Line Color:", 
                value = "#CD2026",
                showColour = "background"
              ),
              numericInput(
                "cor_line_size", 
                "Line Size:", 
                value = 1.5, 
                min = 0.5, 
                max = 5,
                step = 0.5,
                width = "100%"
              )
            ),
            sliderInput(
              "cor_conf", 
              "Confidence Level:", 
              min = 0.90, 
              max = 0.99, 
              value = 0.95, 
              step = 0.01,
              width = "100%"
            ),
            actionButton(
              "run_cor", 
              "Run Correlation Analysis", 
              class = "btn-primary btn-action",
              icon = icon("calculator"),
              width = "100%",
              style = "margin-top: 10px;"
            ),
            downloadButton(
              "download_cor", 
              "Download Results (Word)",
              class = "btn-success btn-action",
              width = "100%",
              style = "margin-top: 10px;"
            ),
            downloadButton(
              "download_cor_plot", 
              "Download Plot",
              class = "btn-info btn-action",
              width = "100%",
              style = "margin-top: 10px;"
            )
          ),
          column(
            8, 
            class = "main-panel",
            withSpinner(
              verbatimTextOutput("cor_results"), 
              type = 4, 
              color = "#0C5EA8",
              size = 1.5
            ),
            conditionalPanel(
              condition = "input.cor_scatter == true",
              div(
                class = "plot-container",
                withSpinner(
                  plotlyOutput("cor_plot"), 
                  type = 4, 
                  color = "#0C5EA8",
                  size = 1.5
                )
              )
            ),
            withSpinner(
              DTOutput("cor_table"), 
              type = 4, 
              color = "#0C5EA8",
              size = 1.5
            )
          )
        )
      ),
      
      tabPanel(
        "ANOVA",
        tabsetPanel(
          tabPanel(
            "One-Way ANOVA",
            fluidRow(
              column(
                4, 
                class = "sidebar-panel",
                div(
                  class = "section-header",
                  h4("One-Way ANOVA Settings", style = "margin: 0;")
                ),
                selectInput(
                  "anova_outcome", 
                  "Outcome Variable (Continuous):", 
                  choices = NULL,
                  width = "100%"
                ),
                selectInput(
                  "anova_group", 
                  "Grouping Variable (Categorical):", 
                  choices = NULL,
                  width = "100%"
                ),
                checkboxInput(
                  "anova_assumptions", 
                  "Check ANOVA Assumptions", 
                  TRUE,
                  width = "100%"
                ),
                checkboxInput(
                  "anova_posthoc", 
                  "Perform Post-hoc Comparisons", 
                  TRUE,
                  width = "100%"
                ),
                conditionalPanel(
                  condition = "input.anova_posthoc",
                  selectInput(
                    "anova_posthoc_method", 
                    "Post-hoc Method:",
                    choices = c(
                      "Tukey HSD" = "tukey",
                      "Bonferroni" = "bonferroni",
                      "Scheffe" = "scheffe",
                      "Dunnett" = "dunnett"
                    ),
                    selected = "tukey",
                    width = "100%"
                  )
                ),
                actionButton(
                  "run_anova", 
                  "Run ANOVA", 
                  class = "btn-primary btn-action",
                  icon = icon("play"),
                  width = "100%",
                  style = "margin-top: 10px;"
                ),
                downloadButton(
                  "download_anova", 
                  "Download Results",
                  class = "btn-success btn-action",
                  width = "100%",
                  style = "margin-top: 10px;"
                )
              ),
              column(
                8, 
                class = "main-panel",
                withSpinner(
                  verbatimTextOutput("anova_results"), 
                  type = 4, 
                  color = "#0C5EA8",
                  size = 1.5
                ),
                withSpinner(
                  DTOutput("anova_table"), 
                  type = 4, 
                  color = "#0C5EA8",
                  size = 1.5
                ),
                conditionalPanel(
                  condition = "input.anova_assumptions",
                  h4("ANOVA Assumptions Check"),
                  withSpinner(
                    verbatimTextOutput("anova_assumptions_results"), 
                    type = 4, 
                    color = "#0C5EA8",
                    size = 1.5
                  ),
                  withSpinner(
                    plotOutput("anova_assumptions_plots"), 
                    type = 4, 
                    color = "#0C5EA8",
                    size = 1.5
                  )
                ),
                conditionalPanel(
                  condition = "input.anova_posthoc",
                  h4("Post-hoc Comparisons"),
                  withSpinner(
                    verbatimTextOutput("anova_posthoc_results"), 
                    type = 4, 
                    color = "#0C5EA8",
                    size = 1.5
                  ),
                  withSpinner(
                    DTOutput("anova_posthoc_table"), 
                    type = 4, 
                    color = "#0C5EA8",
                    size = 1.5
                  ),
                  withSpinner(
                    plotOutput("anova_posthoc_plot"), 
                    type = 4, 
                    color = "#0C5EA8",
                    size = 1.5
                  )
                )
              )
            )
          ),
          tabPanel(
            "Two-Way ANOVA",
            fluidRow(
              column(
                4, 
                class = "sidebar-panel",
                div(
                  class = "section-header",
                  h4("Two-Way ANOVA Settings", style = "margin: 0;")
                ),
                selectInput(
                  "anova_outcome_2way", 
                  "Outcome Variable:", 
                  choices = NULL,
                  width = "100%"
                ),
                selectInput(
                  "anova_factor1", 
                  "Factor 1:", 
                  choices = NULL,
                  width = "100%"
                ),
                selectInput(
                  "anova_factor2", 
                  "Factor 2:", 
                  choices = NULL,
                  width = "100%"
                ),
                checkboxInput(
                  "anova_interaction", 
                  "Include Interaction Term", 
                  TRUE,
                  width = "100%"
                ),
                checkboxInput(
                  "anova_2way_posthoc", 
                  "Perform Post-hoc Tests", 
                  TRUE,
                  width = "100%"
                ),
                actionButton(
                  "run_anova_2way", 
                  "Run Two-Way ANOVA", 
                  class = "btn-primary btn-action",
                  icon = icon("play"),
                  width = "100%",
                  style = "margin-top: 10px;"
                ),
                downloadButton(
                  "download_anova_2way", 
                  "Download Results",
                  class = "btn-success btn-action",
                  width = "100%",
                  style = "margin-top: 10px;"
                )
              ),
              column(
                8, 
                class = "main-panel",
                withSpinner(
                  verbatimTextOutput("anova_2way_results"), 
                  type = 4, 
                  color = "#0C5EA8",
                  size = 1.5
                ),
                withSpinner(
                  DTOutput("anova_2way_table"), 
                  type = 4, 
                  color = "#0C5EA8",
                  size = 1.5
                ),
                conditionalPanel(
                  condition = "input.anova_2way_posthoc",
                  h4("Simple Effects Analysis"),
                  withSpinner(
                    verbatimTextOutput("anova_2way_posthoc"), 
                    type = 4, 
                    color = "#0C5EA8",
                    size = 1.5
                  )
                )
              )
            )
          ),
          tabPanel(
            "Welch ANOVA",
            fluidRow(
              column(
                4, 
                class = "sidebar-panel",
                div(
                  class = "section-header",
                  h4("Welch ANOVA Settings", style = "margin: 0;")
                ),
                selectInput(
                  "welch_outcome", 
                  "Outcome Variable:", 
                  choices = NULL,
                  width = "100%"
                ),
                selectInput(
                  "welch_group", 
                  "Grouping Variable:", 
                  choices = NULL,
                  width = "100%"
                ),
                helpText(
                  "Welch ANOVA is robust to violations of homogeneity of variance."
                ),
                actionButton(
                  "run_welch_anova", 
                  "Run Welch ANOVA", 
                  class = "btn-primary btn-action",
                  icon = icon("play"),
                  width = "100%",
                  style = "margin-top: 10px;"
                ),
                downloadButton(
                  "download_welch_anova", 
                  "Download Results",
                  class = "btn-success btn-action",
                  width = "100%",
                  style = "margin-top: 10px;"
                )
              ),
              column(
                8, 
                class = "main-panel",
                withSpinner(
                  verbatimTextOutput("welch_anova_results"), 
                  type = 4, 
                  color = "#0C5EA8",
                  size = 1.5
                ),
                withSpinner(
                  DTOutput("welch_anova_table"), 
                  type = 4, 
                  color = "#0C5EA8",
                  size = 1.5
                ),
                h4("Games-Howell Post-hoc Test"),
                withSpinner(
                  verbatimTextOutput("welch_posthoc_results"), 
                  type = 4, 
                  color = "#0C5EA8",
                  size = 1.5
                )
              )
            )
          ),
          tabPanel(
            "ANOVA Assumptions",
            fluidRow(
              column(
                4, 
                class = "sidebar-panel",
                div(
                  class = "section-header",
                  h4("ANOVA Assumptions Check", style = "margin: 0;")
                ),
                selectInput(
                  "assumptions_outcome", 
                  "Outcome Variable:", 
                  choices = NULL,
                  width = "100%"
                ),
                selectInput(
                  "assumptions_group", 
                  "Grouping Variable:", 
                  choices = NULL,
                  width = "100%"
                ),
                h4("Tests for Assumptions"),
                checkboxInput(
                  "levene_test", 
                  "Levene's Test (Homogeneity of Variance)", 
                  TRUE,
                  width = "100%"
                ),
                checkboxInput(
                  "normality_test", 
                  "Normality Tests", 
                  TRUE,
                  width = "100%"
                ),
                checkboxInput(
                  "outlier_test", 
                  "Outlier Detection", 
                  TRUE,
                  width = "100%"
                ),
                actionButton(
                  "run_assumptions", 
                  "Check Assumptions", 
                  class = "btn-primary btn-action",
                  icon = icon("check"),
                  width = "100%",
                  style = "margin-top: 10px;"
                ),
                downloadButton(
                  "download_assumptions", 
                  "Download Report",
                  class = "btn-success btn-action",
                  width = "100%",
                  style = "margin-top: 10px;"
                )
              ),
              column(
                8, 
                class = "main-panel",
                conditionalPanel(
                  condition = "input.levene_test",
                  h4("Homogeneity of Variance Tests"),
                  withSpinner(
                    verbatimTextOutput("levene_results"), 
                    type = 4, 
                    color = "#0C5EA8",
                    size = 1.5
                  )
                ),
                conditionalPanel(
                  condition = "input.normality_test",
                  h4("Normality Tests"),
                  withSpinner(
                    verbatimTextOutput("normality_results"), 
                    type = 4, 
                    color = "#0C5EA8",
                    size = 1.5
                  )
                ),
                conditionalPanel(
                  condition = "input.outlier_test",
                  h4("Outlier Detection"),
                  withSpinner(
                    verbatimTextOutput("outlier_results"), 
                    type = 4, 
                    color = "#0C5EA8",
                    size = 1.5
                  )
                ),
                h4("Diagnostic Plots"),
                withSpinner(
                  plotOutput("assumptions_plots"), 
                  type = 4, 
                  color = "#0C5EA8",
                  size = 1.5
                )
              )
            )
          )
        )
      ),
      tabPanel(
        "T-Tests",
        tabsetPanel(
          tabPanel(
            "One-Sample T-Test",
            fluidRow(
              column(
                4, 
                class = "sidebar-panel",
                div(
                  class = "section-header",
                  h4("One-Sample T-Test Settings", style = "margin: 0;")
                ),
                selectInput(
                  "ttest_onesample_var", 
                  "Variable:", 
                  choices = NULL,
                  width = "100%"
                ),
                numericInput(
                  "ttest_onesample_mu", 
                  "Test Value (μ₀):", 
                  value = 0,
                  width = "100%"
                ),
                selectInput(
                  "ttest_onesample_alternative", 
                  "Alternative Hypothesis:",
                  choices = c(
                    "Two-sided" = "two.sided",
                    "Greater than" = "greater",
                    "Less than" = "less"
                  ),
                  selected = "two.sided",
                  width = "100%"
                ),
                sliderInput(
                  "ttest_onesample_conf", 
                  "Confidence Level:", 
                  min = 0.90, 
                  max = 0.99, 
                  value = 0.95, 
                  step = 0.01,
                  width = "100%"
                ),
                actionButton(
                  "run_ttest_onesample", 
                  "Run One-Sample T-Test", 
                  class = "btn-primary btn-action",
                  icon = icon("play"),
                  width = "100%",
                  style = "margin-top: 10px;"
                ),
                downloadButton(
                  "download_ttest_onesample", 
                  "Download Results",
                  class = "btn-success btn-action",
                  width = "100%",
                  style = "margin-top: 10px;"
                )
              ),
              column(
                8, 
                class = "main-panel",
                withSpinner(
                  verbatimTextOutput("ttest_onesample_results"), 
                  type = 4, 
                  color = "#0C5EA8",
                  size = 1.5
                ),
                withSpinner(
                  DTOutput("ttest_onesample_table"), 
                  type = 4, 
                  color = "#0C5EA8",
                  size = 1.5
                ),
                h4("Descriptive Statistics"),
                withSpinner(
                  verbatimTextOutput("ttest_onesample_descriptives"), 
                  type = 4, 
                  color = "#0C5EA8",
                  size = 1.5
                ),
                withSpinner(
                  plotOutput("ttest_onesample_plot"), 
                  type = 4, 
                  color = "#0C5EA8",
                  size = 1.5
                )
              )
            )
          ),
          tabPanel(
            "Independent T-Test",
            fluidRow(
              column(
                4, 
                class = "sidebar-panel",
                div(
                  class = "section-header",
                  h4("Independent T-Test Settings", style = "margin: 0;")
                ),
                selectInput(
                  "ttest_independent_var", 
                  "Outcome Variable:", 
                  choices = NULL,
                  width = "100%"
                ),
                selectInput(
                  "ttest_independent_group", 
                  "Grouping Variable:", 
                  choices = NULL,
                  width = "100%"
                ),
                uiOutput("ttest_independent_levels_ui"),
                checkboxInput(
                  "ttest_var_equal", 
                  "Assume Equal Variances", 
                  FALSE,
                  width = "100%"
                ),
                selectInput(
                  "ttest_alternative", 
                  "Alternative Hypothesis:",
                  choices = c(
                    "Two-sided" = "two.sided",
                    "Greater than" = "greater",
                    "Less than" = "less"
                  ),
                  selected = "two.sided",
                  width = "100%"
                ),
                sliderInput(
                  "ttest_conf", 
                  "Confidence Level:", 
                  min = 0.90, 
                  max = 0.99, 
                  value = 0.95, 
                  step = 0.01,
                  width = "100%"
                ),
                actionButton(
                  "run_ttest_independent", 
                  "Run Independent T-Test", 
                  class = "btn-primary btn-action",
                  icon = icon("play"),
                  width = "100%",
                  style = "margin-top: 10px;"
                ),
                downloadButton(
                  "download_ttest_independent", 
                  "Download Results",
                  class = "btn-success btn-action",
                  width = "100%",
                  style = "margin-top: 10px;"
                )
              ),
              column(
                8, 
                class = "main-panel",
                withSpinner(
                  verbatimTextOutput("ttest_independent_results"), 
                  type = 4, 
                  color = "#0C5EA8",
                  size = 1.5
                ),
                withSpinner(
                  DTOutput("ttest_independent_table"), 
                  type = 4, 
                  color = "#0C5EA8",
                  size = 1.5
                ),
                h4("Group Descriptives"),
                withSpinner(
                  verbatimTextOutput("ttest_independent_descriptives"), 
                  type = 4, 
                  color = "#0C5EA8",
                  size = 1.5
                ),
                h4("Variance Check"),
                withSpinner(
                  verbatimTextOutput("ttest_variance_check"), 
                  type = 4, 
                  color = "#0C5EA8",
                  size = 1.5
                ),
                withSpinner(
                  plotOutput("ttest_independent_plot"), 
                  type = 4, 
                  color = "#0C5EA8",
                  size = 1.5
                )
              )
            )
          ),
          tabPanel(
            "Paired T-Test",
            fluidRow(
              column(
                4, 
                class = "sidebar-panel",
                div(
                  class = "section-header",
                  h4("Paired T-Test Settings", style = "margin: 0;")
                ),
                selectInput(
                  "ttest_paired_var1", 
                  "Variable 1 (Before/Time 1):", 
                  choices = NULL,
                  width = "100%"
                ),
                selectInput(
                  "ttest_paired_var2", 
                  "Variable 2 (After/Time 2):", 
                  choices = NULL,
                  width = "100%"
                ),
                selectInput(
                  "ttest_paired_alternative", 
                  "Alternative Hypothesis:",
                  choices = c(
                    "Two-sided" = "two.sided",
                    "Greater than" = "greater",
                    "Less than" = "less"
                  ),
                  selected = "two.sided",
                  width = "100%"
                ),
                sliderInput(
                  "ttest_paired_conf", 
                  "Confidence Level:", 
                  min = 0.90, 
                  max = 0.99, 
                  value = 0.95, 
                  step = 0.01,
                  width = "100%"
                ),
                actionButton(
                  "run_ttest_paired", 
                  "Run Paired T-Test", 
                  class = "btn-primary btn-action",
                  icon = icon("play"),
                  width = "100%",
                  style = "margin-top: 10px;"
                ),
                downloadButton(
                  "download_ttest_paired", 
                  "Download Results",
                  class = "btn-success btn-action",
                  width = "100%",
                  style = "margin-top: 10px;"
                )
              ),
              column(
                8, 
                class = "main-panel",
                withSpinner(
                  verbatimTextOutput("ttest_paired_results"), 
                  type = 4, 
                  color = "#0C5EA8",
                  size = 1.5
                ),
                withSpinner(
                  DTOutput("ttest_paired_table"), 
                  type = 4, 
                  color = "#0C5EA8",
                  size = 1.5
                ),
                h4("Paired Descriptives"),
                withSpinner(
                  verbatimTextOutput("ttest_paired_descriptives"), 
                  type = 4, 
                  color = "#0C5EA8",
                  size = 1.5
                ),
                h4("Difference Analysis"),
                withSpinner(
                  verbatimTextOutput("ttest_paired_differences"), 
                  type = 4, 
                  color = "#0C5EA8",
                  size = 1.5
                ),
                withSpinner(
                  plotOutput("ttest_paired_plot"), 
                  type = 4, 
                  color = "#0C5EA8",
                  size = 1.5
                )
              )
            )
          ),
          tabPanel(
            "Non-parametric Tests",
            fluidRow(
              column(
                4, 
                class = "sidebar-panel",
                div(
                  class = "section-header",
                  h4("Non-parametric Tests Settings", style = "margin: 0;")
                ),
                h4("Mann-Whitney U Test"),
                selectInput(
                  "mannwhitney_var", 
                  "Outcome Variable:", 
                  choices = NULL,
                  width = "100%"
                ),
                selectInput(
                  "mannwhitney_group", 
                  "Grouping Variable:", 
                  choices = NULL,
                  width = "100%"
                ),
                actionButton(
                  "run_mannwhitney", 
                  "Run Mann-Whitney Test", 
                  class = "btn-primary btn-action",
                  icon = icon("play"),
                  width = "100%",
                  style = "margin-top: 10px;"
                ),
                tags$hr(),
                h4("Wilcoxon Signed-Rank Test"),
                selectInput(
                  "wilcoxon_var1", 
                  "Variable 1:", 
                  choices = NULL,
                  width = "100%"
                ),
                selectInput(
                  "wilcoxon_var2", 
                  "Variable 2:", 
                  choices = NULL,
                  width = "100%"
                ),
                actionButton(
                  "run_wilcoxon", 
                  "Run Wilcoxon Test", 
                  class = "btn-primary btn-action",
                  icon = icon("play"),
                  width = "100%",
                  style = "margin-top: 10px;"
                ),
                downloadButton(
                  "download_nonparametric", 
                  "Download Results",
                  class = "btn-success btn-action",
                  width = "100%",
                  style = "margin-top: 20px;"
                )
              ),
              column(
                8, 
                class = "main-panel",
                h4("Mann-Whitney U Test Results"),
                withSpinner(
                  verbatimTextOutput("mannwhitney_results"), 
                  type = 4, 
                  color = "#0C5EA8",
                  size = 1.5
                ),
                tags$hr(),
                h4("Wilcoxon Signed-Rank Test Results"),
                withSpinner(
                  verbatimTextOutput("wilcoxon_results"), 
                  type = 4, 
                  color = "#0C5EA8",
                  size = 1.5
                )
              )
            )
          )
        )
      )
    )
  )
)

# Server function
server <- function(input, output, session) {
  
  # Add loading states for all action buttons
  observeEvent(input$load_data, {
    shinyjs::disable("load_data")
    shinyjs::html("load_data", "<i class='fas fa-spinner fa-spin'></i> Loading...")
  })
  
  observeEvent(input$run_freq, {
    shinyjs::disable("run_freq")
    shinyjs::html("run_freq", "<i class='fas fa-spinner fa-spin'></i> Analyzing...")
  })
  
  observeEvent(input$run_cor, {
    shinyjs::disable("run_cor")
    shinyjs::html("run_cor", "<i class='fas fa-spinner fa-spin'></i> Calculating...")
  })
  
  observeEvent(input$run_poisson, {
    shinyjs::disable("run_poisson")
    shinyjs::html("run_poisson", "<i class='fas fa-spinner fa-spin'></i> Running...")
  })
  
  observeEvent(input$run_central, {
    shinyjs::disable("run_central")
    shinyjs::html("run_central", "<i class='fas fa-spinner fa-spin'></i> Analyzing...")
  })
  
  observeEvent(input$run_epi, {
    shinyjs::disable("run_epi")
    shinyjs::html("run_epi", "<i class='fas fa-spinner fa-spin'></i> Generating...")
  })
  
  observeEvent(input$run_ip, {
    shinyjs::disable("run_ip")
    shinyjs::html("run_ip", "<i class='fas fa-spinner fa-spin'></i> Calculating...")
  })
  
  observeEvent(input$run_rr, {
    shinyjs::disable("run_rr")
    shinyjs::html("run_rr", "<i class='fas fa-spinner fa-spin'></i> Calculating...")
  })
  
  observeEvent(input$run_or, {
    shinyjs::disable("run_or")
    shinyjs::html("run_or", "<i class='fas fa-spinner fa-spin'></i> Calculating...")
  })
  
  observeEvent(input$run_surv, {
    shinyjs::disable("run_surv")
    shinyjs::html("run_surv", "<i class='fas fa-spinner fa-spin'></i> Analyzing...")
  })
  
  observeEvent(input$run_cox, {
    shinyjs::disable("run_cox")
    shinyjs::html("run_cox", "<i class='fas fa-spinner fa-spin'></i> Running...")
  })
  
  observeEvent(input$run_linear, {
    shinyjs::disable("run_linear")
    shinyjs::html("run_linear", "<i class='fas fa-spinner fa-spin'></i> Running...")
  })
  
  observeEvent(input$run_logistic, {
    shinyjs::disable("run_logistic")
    shinyjs::html("run_logistic", "<i class='fas fa-spinner fa-spin'></i> Running...")
  })
  
  observeEvent(input$run_chisq, {
    shinyjs::disable("run_chisq")
    shinyjs::html("run_chisq", "<i class='fas fa-spinner fa-spin'></i> Testing...")
  })
  
  observeEvent(input$run_anova, {
    shinyjs::disable("run_anova")
    shinyjs::html("run_anova", "<i class='fas fa-spinner fa-spin'></i> Running...")
  })
  
  observeEvent(input$run_ttest_onesample, {
    shinyjs::disable("run_ttest_onesample")
    shinyjs::html("run_ttest_onesample", "<i class='fas fa-spinner fa-spin'></i> Running...")
  })
  
  # Reset button states after calculations
  observe({
    if (!is.null(processed_data())) {
      shinyjs::enable("load_data")
      shinyjs::html("load_data", "<i class='fas fa-upload'></i> Load Data")
    }
  })
  
  # Reset analysis buttons after their respective calculations complete
  observeEvent(freq_results(), {
    shinyjs::enable("run_freq")
    shinyjs::html("run_freq", "<i class='fas fa-play'></i> Run Analysis")
  })
  
  observeEvent(cor_results(), {
    shinyjs::enable("run_cor")
    shinyjs::html("run_cor", "<i class='fas fa-calculator'></i> Run Correlation Analysis")
  })
  
  observeEvent(central_results(), {
    shinyjs::enable("run_central")
    shinyjs::html("run_central", "<i class='fas fa-play'></i> Run Analysis")
  })
  
  observeEvent(epi_results(), {
    shinyjs::enable("run_epi")
    shinyjs::html("run_epi", "<i class='fas fa-play'></i> Run Analysis")
  })
  
  observeEvent(ip_results(), {
    shinyjs::enable("run_ip")
    shinyjs::html("run_ip", "<i class='fas fa-calculator'></i> Calculate")
  })
  
  observeEvent(rr_results(), {
    shinyjs::enable("run_rr")
    shinyjs::html("run_rr", "<i class='fas fa-calculator'></i> Calculate")
  })
  
  observeEvent(or_results(), {
    shinyjs::enable("run_or")
    shinyjs::html("run_or", "<i class='fas fa-calculator'></i> Calculate")
  })
  
  observeEvent(surv_results(), {
    shinyjs::enable("run_surv")
    shinyjs::html("run_surv", "<i class='fas fa-play'></i> Run Analysis")
  })
  
  observeEvent(cox_results(), {
    shinyjs::enable("run_cox")
    shinyjs::html("run_cox", "<i class='fas fa-play'></i> Run Analysis")
  })
  
  observeEvent(linear_results(), {
    shinyjs::enable("run_linear")
    shinyjs::html("run_linear", "<i class='fas fa-play'></i> Run Analysis")
  })
  
  observeEvent(logistic_results(), {
    shinyjs::enable("run_logistic")
    shinyjs::html("run_logistic", "<i class='fas fa-play'></i> Run Analysis")
  })
  
  observeEvent(chisq_results(), {
    shinyjs::enable("run_chisq")
    shinyjs::html("run_chisq", "<i class='fas fa-play'></i> Run Test")
  })
  
  observeEvent(anova_results(), {
    shinyjs::enable("run_anova")
    shinyjs::html("run_anova", "<i class='fas fa-play'></i> Run ANOVA")
  })
  
  observeEvent(ttest_onesample_results(), {
    shinyjs::enable("run_ttest_onesample")
    shinyjs::html("run_ttest_onesample", "<i class='fas fa-play'></i> Run One-Sample T-Test")
  })
  
  # Reactive data with proper handling
  raw_data <- reactiveVal()
  processed_data <- reactiveVal()
  
  # Load data with proper error handling
  observeEvent(input$load_data, {
    req(input$file1)
    
    tryCatch({
      ext <- tools::file_ext(input$file1$name)
      
      df <- switch(ext,
                   csv = read.csv(input$file1$datapath,
                                  header = input$header,
                                  sep = input$sep,
                                  quote = input$quote,
                                  stringsAsFactors = FALSE),
                   xls = read_excel(input$file1$datapath, sheet = input$sheet),
                   xlsx = read_excel(input$file1$datapath, sheet = input$sheet),
                   sav = read_sav(input$file1$datapath),
                   dta = read_dta(input$file1$datapath),
                   sas7bdat = haven::read_sas(input$file1$datapath),
                   RData = {
                     env <- new.env()
                     load(input$file1$datapath, env)
                     get(ls(env)[1], env)
                   },
                   rda = {
                     env <- new.env()
                     load(input$file1$datapath, env)
                     get(ls(env)[1], env)
                   },
                   validate("Invalid file format")
      )
      
      # Convert to data.frame and handle labelled data
      if (inherits(df, "tbl_df")) {
        df <- as.data.frame(df)
      }
      
      # Handle labelled data from SPSS/Stata
      df <- convert_labelled(df)
      
      raw_data(df)
      
      # Process data - remove missing values
      df_processed <- df %>%
        mutate(across(everything(), ~ifelse(is.na(.), NA, .))) %>%
        na.omit()
      
      processed_data(df_processed)
      
      # Update all select inputs
      updateSelectInput(session, "freq_var", choices = names(df_processed))
      updateSelectInput(session, "central_var", choices = names(df_processed)[sapply(df_processed, is.numeric)])
      updateSelectInput(session, "epi_date", choices = names(df_processed))
      updateSelectInput(session, "epi_case", choices = names(df_processed))
      updateSelectInput(session, "epi_group", choices = c("None" = "none", names(df_processed)))
      updateSelectInput(session, "ip_case", choices = names(df_processed))
      updateSelectInput(session, "ip_pop", choices = names(df_processed)[sapply(df_processed, is.numeric)])
      updateSelectInput(session, "ip_time", choices = names(df_processed))
      updateSelectInput(session, "rr_outcome", choices = names(df_processed))
      updateSelectInput(session, "rr_exposure", choices = names(df_processed))
      updateSelectInput(session, "or_outcome", choices = names(df_processed))
      updateSelectInput(session, "or_exposure", choices = names(df_processed))
      updateSelectInput(session, "surv_time", choices = names(df_processed)[sapply(df_processed, is.numeric)])
      updateSelectInput(session, "surv_event", choices = names(df_processed))
      updateSelectInput(session, "surv_group", choices = c("None" = "none", names(df_processed)))
      updateSelectInput(session, "cox_time", choices = names(df_processed)[sapply(df_processed, is.numeric)])
      updateSelectInput(session, "cox_event", choices = names(df_processed))
      updateSelectInput(session, "cox_vars", choices = names(df_processed))
      updateSelectInput(session, "linear_outcome", choices = names(df_processed)[sapply(df_processed, is.numeric)])
      updateSelectInput(session, "linear_vars", choices = names(df_processed))
      updateSelectInput(session, "logistic_outcome", choices = names(df_processed))
      updateSelectInput(session, "logistic_vars", choices = names(df_processed))
      updateSelectInput(session, "chisq_var1", choices = names(df_processed))
      updateSelectInput(session, "chisq_var2", choices = names(df_processed))
      
      # Update correlation analysis inputs
      updateSelectInput(session, "cor_var1", choices = names(df_processed)[sapply(df_processed, is.numeric)])
      updateSelectInput(session, "cor_var2", choices = names(df_processed)[sapply(df_processed, is.numeric)])
      
      # Update ANOVA inputs
      updateSelectInput(session, "anova_outcome", choices = names(df_processed)[sapply(df_processed, is.numeric)])
      updateSelectInput(session, "anova_group", choices = names(df_processed))
      updateSelectInput(session, "anova_outcome_2way", choices = names(df_processed)[sapply(df_processed, is.numeric)])
      updateSelectInput(session, "anova_factor1", choices = names(df_processed))
      updateSelectInput(session, "anova_factor2", choices = names(df_processed))
      updateSelectInput(session, "welch_outcome", choices = names(df_processed)[sapply(df_processed, is.numeric)])
      updateSelectInput(session, "welch_group", choices = names(df_processed))
      updateSelectInput(session, "assumptions_outcome", choices = names(df_processed)[sapply(df_processed, is.numeric)])
      updateSelectInput(session, "assumptions_group", choices = names(df_processed))
      
      # Update T-test inputs
      updateSelectInput(session, "ttest_onesample_var", choices = names(df_processed)[sapply(df_processed, is.numeric)])
      updateSelectInput(session, "ttest_independent_var", choices = names(df_processed)[sapply(df_processed, is.numeric)])
      updateSelectInput(session, "ttest_independent_group", choices = names(df_processed))
      updateSelectInput(session, "ttest_paired_var1", choices = names(df_processed)[sapply(df_processed, is.numeric)])
      updateSelectInput(session, "ttest_paired_var2", choices = names(df_processed)[sapply(df_processed, is.numeric)])
      updateSelectInput(session, "mannwhitney_var", choices = names(df_processed)[sapply(df_processed, is.numeric)])
      updateSelectInput(session, "mannwhitney_group", choices = names(df_processed))
      updateSelectInput(session, "wilcoxon_var1", choices = names(df_processed)[sapply(df_processed, is.numeric)])
      updateSelectInput(session, "wilcoxon_var2", choices = names(df_processed)[sapply(df_processed, is.numeric)])
      
      # Update Poisson regression inputs
      updateSelectInput(session, "poisson_outcome", choices = names(df_processed)[sapply(df_processed, is.numeric)])
      updateSelectInput(session, "poisson_vars", choices = names(df_processed))
      updateSelectInput(session, "poisson_offset_var", choices = names(df_processed)[sapply(df_processed, is.numeric)])
      
      # Update central tendency inputs
      updateSelectInput(session, "central_stratify", choices = c("None" = "none", names(df_processed)))
      
      # Show success notification
      showNotification("Data loaded successfully!", type = "message", duration = 3)
      
    }, error = function(e) {
      showNotification(paste("Error loading file:", e$message), type = "error")
    })
  })
  
  # Missing values alert
  output$missing_values_alert <- renderUI({
    req(raw_data())
    df <- raw_data()
    missing_count <- sum(is.na(df))
    total_cells <- nrow(df) * ncol(df)
    missing_pct <- round((missing_count / total_cells) * 100, 1)
    
    if (missing_count > 0) {
      div(
        class = "alert alert-warning",
        h4("Missing Values Detected"),
        p(paste("Found", missing_count, "missing values (", missing_pct, "%) in the dataset.")),
        p("Missing values will be automatically excluded from all analyses.")
      )
    } else {
      div(
        class = "alert alert-success",
        h4("No Missing Values"),
        p("All data points are complete.")
      )
    }
  })
  
  # Data preview
  output$contents <- renderDT({
    req(processed_data())
    datatable(processed_data(), 
              options = list(
                scrollX = TRUE, 
                pageLength = 10,
                dom = 'Blfrtip',
                buttons = c('copy', 'csv', 'excel', 'pdf')
              ),
              rownames = FALSE,
              extensions = 'Buttons',
              class = 'cell-border stripe hover'
    ) %>%
      formatStyle(names(processed_data()),
                  backgroundColor = 'white',
                  fontWeight = 'normal')
  })
  
  # Data summary
  output$data_summary <- renderPrint({
    req(processed_data())
    df <- processed_data()
    cat("Dataset Summary:\n")
    cat("Number of observations:", nrow(df), "\n")
    cat("Number of variables:", ncol(df), "\n\n")
    cat("Variable types:\n")
    type_summary <- table(sapply(df, class))
    for (type in names(type_summary)) {
      cat(paste0("  ", type, ": ", type_summary[type], "\n"))
    }
  })
  
  # Ultra-robust Frequency Analysis
  freq_results <- eventReactive(input$run_freq, {
    req(input$freq_var, processed_data())
    
    tryCatch({
      df <- processed_data()
      
      # Double-check variable exists
      if (!input$freq_var %in% names(df)) {
        return(list(
          table = data.frame(Error = "Variable not found in dataset"),
          success = FALSE
        ))
      }
      
      # Safely get the variable
      var_data <- df[[input$freq_var]]
      
      # Handle empty or all-NA data
      if (length(var_data) == 0 || all(is.na(var_data))) {
        return(list(
          table = data.frame(Error = "Variable contains no data"),
          success = FALSE
        ))
      }
      
      # Simple frequency table using base R (most reliable)
      freq_counts <- table(var_data, useNA = "ifany")
      freq_table <- data.frame(
        Category = names(freq_counts),
        Frequency = as.numeric(freq_counts),
        stringsAsFactors = FALSE
      )
      
      # Calculate proportions and percentages
      total <- sum(freq_table$Frequency)
      freq_table$Proportion <- freq_table$Frequency / total
      freq_table$Percentage <- freq_table$Proportion * 100
      
      # Clean up category names
      freq_table$Category[is.na(freq_table$Category)] <- "Missing"
      
      list(table = freq_table, success = TRUE)
      
    }, error = function(e) {
      return(list(
        table = data.frame(Error = paste("Error:", e$message)),
        success = FALSE
      ))
    })
  })
  
  output$freq_table <- renderDT({
    req(freq_results())
    
    results <- freq_results()
    
    if (!results$success) {
      datatable(results$table, 
                options = list(dom = 't'),
                rownames = FALSE,
                caption = "Frequency Analysis - Error")
    } else {
      datatable(results$table,
                options = list(
                  pageLength = 10,
                  dom = 'Blfrtip',
                  buttons = c('copy', 'csv', 'excel')
                ),
                rownames = FALSE,
                caption = "Frequency Distribution",
                extensions = 'Buttons'
      ) %>%
        formatRound(columns = c('Proportion', 'Percentage'), digits = 2) %>%
        formatStyle(
          'Frequency',
          background = styleColorBar(range(results$table$Frequency), '#0C5EA8'),
          backgroundSize = '100% 90%',
          backgroundRepeat = 'no-repeat',
          backgroundPosition = 'center'
        )
    }
  })
  
  output$freq_plot <- renderPlotly({
    req(freq_results())
    
    results <- freq_results()
    
    if (!results$success) {
      p <- plot_ly() %>%
        add_annotations(
          text = results$table$Error[1],
          x = 0.5,
          y = 0.5,
          xref = "paper",
          yref = "paper",
          showarrow = FALSE,
          font = list(size = 16, color = "red")
        ) %>%
        layout(
          title = "Frequency Plot - Error",
          xaxis = list(showgrid = FALSE, zeroline = FALSE, showticklabels = FALSE),
          yaxis = list(showgrid = FALSE, zeroline = FALSE, showticklabels = FALSE),
          margin = list(t = 50)
        )
      return(p)
    }
    
    # Create bar plot
    p <- plot_ly(results$table,
                 x = ~Category,
                 y = ~Frequency,
                 type = 'bar',
                 marker = list(color = input$freq_color,
                               line = list(color = 'rgba(0,0,0,0.2)', width = 1)),
                 hoverinfo = 'text',
                 text = ~paste('Category:', Category,
                               '<br>Count:', Frequency,
                               '<br>Percentage:', round(Percentage, 1), '%'),
                 textposition = 'none') %>%
      layout(
        title = list(text = input$freq_title, 
                     font = list(size = input$freq_title_size, 
                                 family = 'Arial, sans-serif')),
        xaxis = list(title = input$freq_var, 
                     tickangle = 45,
                     titlefont = list(size = input$freq_x_size)),
        yaxis = list(title = "Frequency",
                     titlefont = list(size = input$freq_y_size)),
        margin = list(b = 100, l = 60, r = 40, t = 60),
        hoverlabel = list(font = list(size = 12)),
        showlegend = FALSE
      )
    
    return(p)
  })
  
 
  # ULTRA-ROBUST CORRELATION ANALYSIS - WITHOUT WGCNA
  cor_results <- eventReactive(input$run_cor, {
    req(input$cor_var1, input$cor_var2, processed_data())
    
    tryCatch({
      df <- processed_data()
      
      # 1. VALIDATE VARIABLES EXIST
      if (!input$cor_var1 %in% names(df) || !input$cor_var2 %in% names(df)) {
        return(list(
          error = "Selected variables not found in dataset",
          success = FALSE
        ))
      }
      
      # 2. GET AND PREPARE VARIABLES
      var1_raw <- df[[input$cor_var1]]
      var2_raw <- df[[input$cor_var2]]
      
      # 3. ULTRA-ROBUST NUMERIC CONVERSION FUNCTION - FIXED VERSION
      convert_to_numeric_safe <- function(x, var_name = "variable") {
        # Store original indices for alignment
        orig_indices <- seq_along(x)
        
        # Remove NA values first and keep track of indices
        non_na_indices <- which(!is.na(x))
        x_clean <- x[non_na_indices]
        orig_indices_clean <- orig_indices[non_na_indices]
        
        if (length(x_clean) == 0) {
          return(list(
            numeric_data = numeric(0),
            indices = integer(0),
            warning = paste(var_name, ": No non-missing values"),
            success = FALSE
          ))
        }
        
        # Try multiple conversion strategies
        conversion_attempts <- list()
        
        # Strategy 1: Direct numeric conversion
        attempt1 <- suppressWarnings(as.numeric(x_clean))
        valid_indices1 <- which(!is.na(attempt1))
        conversion_attempts[[1]] <- list(
          method = "as.numeric()",
          valid_indices = orig_indices_clean[valid_indices1],
          data = attempt1[valid_indices1],
          valid_count = length(valid_indices1)
        )
        
        # Strategy 2: Convert factor to numeric via levels
        if (is.factor(x_clean)) {
          attempt2 <- as.numeric(as.character(x_clean))
          valid_indices2 <- which(!is.na(attempt2))
          conversion_attempts[[2]] <- list(
            method = "as.numeric(as.character()) for factor",
            valid_indices = orig_indices_clean[valid_indices2],
            data = attempt2[valid_indices2],
            valid_count = length(valid_indices2)
          )
        }
        
        # Strategy 3: Character to numeric with pattern matching
        if (is.character(x_clean)) {
          # Extract numbers from strings
          attempt3 <- sapply(x_clean, function(val) {
            # Try direct conversion first
            num_val <- suppressWarnings(as.numeric(val))
            if (!is.na(num_val)) return(num_val)
            
            # Extract first number found in string
            matches <- regmatches(val, gregexpr("[-+]?[0-9]*\\.?[0-9]+([eE][-+]?[0-9]+)?", val))
            if (length(matches[[1]]) > 0) {
              as.numeric(matches[[1]][1])
            } else {
              NA_real_
            }
          })
          valid_indices3 <- which(!is.na(attempt3))
          conversion_attempts[[3]] <- list(
            method = "numeric extraction from character",
            valid_indices = orig_indices_clean[valid_indices3],
            data = attempt3[valid_indices3],
            valid_count = length(valid_indices3)
          )
        }
        
        # Strategy 4: Convert logical to numeric
        if (is.logical(x_clean)) {
          attempt4 <- as.numeric(x_clean)
          valid_indices4 <- which(!is.na(attempt4))
          conversion_attempts[[4]] <- list(
            method = "logical to numeric",
            valid_indices = orig_indices_clean[valid_indices4],
            data = attempt4[valid_indices4],
            valid_count = length(valid_indices4)
          )
        }
        
        # Choose the best conversion method
        best_method <- NULL
        best_data <- NULL
        best_indices <- NULL
        max_valid <- 0
        
        for (attempt in conversion_attempts) {
          if (attempt$valid_count > max_valid && attempt$valid_count > 0) {
            max_valid <- attempt$valid_count
            best_method <- attempt$method
            best_data <- attempt$data
            best_indices <- attempt$valid_indices
          }
        }
        
        if (is.null(best_method)) {
          return(list(
            numeric_data = numeric(0),
            indices = integer(0),
            warning = paste(var_name, ": No valid numeric conversion found"),
            success = FALSE
          ))
        }
        
        # Check if we have enough data
        if (length(best_data) < 3) {
          return(list(
            numeric_data = best_data,
            indices = best_indices,
            warning = paste(var_name, ": Only", length(best_data), "valid numeric values after conversion"),
            success = TRUE  # Still success, but with warning
          ))
        }
        
        return(list(
          numeric_data = best_data,
          indices = best_indices,
          method = best_method,
          original_length = length(x),
          converted_length = length(best_data),
          success = TRUE
        ))
      }
      
      # 4. APPLY CONVERSION TO BOTH VARIABLES
      var1_conversion <- convert_to_numeric_safe(var1_raw, input$cor_var1)
      var2_conversion <- convert_to_numeric_safe(var2_raw, input$cor_var2)
      
      # 5. CHECK CONVERSION SUCCESS
      if (!var1_conversion$success || !var2_conversion$success) {
        error_msg <- paste(
          "Numeric conversion failed:",
          if (!var1_conversion$success) paste("\n-", input$cor_var1, ":", var1_conversion$warning),
          if (!var2_conversion$success) paste("\n-", input$cor_var2, ":", var2_conversion$warning),
          "\n\nPlease ensure variables contain numeric data or can be converted to numeric."
        )
        return(list(
          error = error_msg,
          success = FALSE
        ))
      }
      
      # 6. ALIGN DATA BY COMMON INDICES (CRITICAL FIX)
      # Find common indices that have valid numeric values in BOTH variables
      common_indices <- intersect(var1_conversion$indices, var2_conversion$indices)
      
      if (length(common_indices) < 3) {
        return(list(
          error = paste("Insufficient data for correlation analysis (only", length(common_indices), 
                        "common valid pairs available)"),
          success = FALSE
        ))
      }
      
      # Extract aligned data using common indices
      get_aligned_data <- function(conversion_result, indices) {
        # Match the indices to find positions in the converted data
        positions <- match(indices, conversion_result$indices)
        return(conversion_result$numeric_data[positions])
      }
      
      var1_final <- get_aligned_data(var1_conversion, common_indices)
      var2_final <- get_aligned_data(var2_conversion, common_indices)
      
      # 7. PERFORM CORRELATION WITH PROPER ERROR HANDLING
      cor_test <- NULL
      
      tryCatch({
        # Use pairwise complete observations for all methods
        if (input$cor_method == "pearson") {
          # For Pearson, use standard cor.test but ensure pairwise complete data
          cor_test <- cor.test(var1_final, var2_final,
                               method = "pearson",
                               conf.level = input$cor_conf,
                               use = "complete.obs")
          
        } else if (input$cor_method == "kendall") {
          # For Kendall, use standard cor function (NO WGCNA)
          cor_value <- cor(var1_final, var2_final, 
                           method = "kendall", 
                           use = "pairwise.complete.obs")
          
          # Calculate p-value and confidence interval
          n <- sum(!is.na(var1_final) & !is.na(var2_final))
          if (n > 3) {
            # Using approximate method for Kendall's tau
            z <- cor_value * sqrt((9 * n * (n - 1)) / (2 * (2 * n + 5)))
            p_value <- 2 * pnorm(-abs(z))
            
            # Approximate confidence interval for Kendall's tau
            se <- sqrt((4 * n + 10) / (9 * n * (n - 1)))
            z_crit <- qnorm(1 - (1 - input$cor_conf) / 2)
            ci_lower <- cor_value - z_crit * se
            ci_upper <- cor_value + z_crit * se
            
            # Ensure bounds are within [-1, 1]
            ci_lower <- max(-1, ci_lower)
            ci_upper <- min(1, ci_upper)
          } else {
            p_value <- NA
            ci_lower <- NA
            ci_upper <- NA
          }
          
          cor_test <- list(
            estimate = cor_value,
            statistic = z,
            p.value = p_value,
            conf.int = c(ci_lower, ci_upper),
            parameter = n - 2,
            alternative = "two.sided",
            method = "Kendall's rank correlation tau",
            data.name = paste(input$cor_var1, "and", input$cor_var2)
          )
          class(cor_test) <- "htest"
          
        } else if (input$cor_method == "spearman") {
          # For Spearman, use standard cor function
          cor_value <- cor(var1_final, var2_final, 
                           method = "spearman", 
                           use = "pairwise.complete.obs")
          
          # Calculate p-value and confidence interval
          n <- sum(!is.na(var1_final) & !is.na(var2_final))
          if (n > 3) {
            # Using approximate method for Spearman's rho
            t_stat <- cor_value * sqrt((n - 2) / (1 - cor_value^2))
            df <- n - 2
            p_value <- 2 * pt(-abs(t_stat), df)
            
            # Confidence interval using Fisher's z transformation
            if (abs(cor_value) < 1) {
              z <- 0.5 * log((1 + cor_value) / (1 - cor_value))
              se <- 1 / sqrt(n - 3)
              z_crit <- qnorm(1 - (1 - input$cor_conf) / 2)
              ci_lower_z <- z - z_crit * se
              ci_upper_z <- z + z_crit * se
              ci_lower <- (exp(2 * ci_lower_z) - 1) / (exp(2 * ci_lower_z) + 1)
              ci_upper <- (exp(2 * ci_upper_z) - 1) / (exp(2 * ci_upper_z) + 1)
            } else {
              ci_lower <- NA
              ci_upper <- NA
            }
          } else {
            p_value <- NA
            ci_lower <- NA
            ci_upper <- NA
          }
          
          cor_test <- list(
            estimate = cor_value,
            statistic = t_stat,
            p.value = p_value,
            conf.int = c(ci_lower, ci_upper),
            parameter = df,
            alternative = "two.sided",
            method = "Spearman's rank correlation rho",
            data.name = paste(input$cor_var1, "and", input$cor_var2)
          )
          class(cor_test) <- "htest"
        }
        
      }, error = function(e) {
        # Fallback to psych package if above fails
        tryCatch({
          if (input$cor_method == "pearson") {
            cor_test <- cor.test(var1_final, var2_final, 
                                 method = "pearson", 
                                 conf.level = input$cor_conf)
          } else {
            # Use psych package for robust correlation
            cor_result <- psych::corr.test(var1_final, var2_final, 
                                           method = input$cor_method,
                                           ci = TRUE, 
                                           conf.level = input$cor_conf)
            
            cor_test <- list(
              estimate = cor_result$r[1, 2],
              p.value = cor_result$p[1, 2],
              conf.int = c(cor_result$ci$lower[1], cor_result$ci$upper[1]),
              statistic = NA,
              parameter = cor_result$n[1, 2] - 2,
              alternative = "two.sided",
              method = paste(tools::toTitleCase(input$cor_method), "correlation (psych package)"),
              data.name = paste(input$cor_var1, "and", input$cor_var2)
            )
            class(cor_test) <- "htest"
          }
        }, error = function(e2) {
          # Ultimate fallback - simple correlation with no inference
          cor_value <- cor(var1_final, var2_final, 
                           method = input$cor_method, 
                           use = "pairwise.complete.obs")
          n <- sum(!is.na(var1_final) & !is.na(var2_final))
          
          cor_test <- list(
            estimate = cor_value,
            statistic = NA,
            p.value = NA,
            conf.int = c(NA, NA),
            parameter = n - 2,
            alternative = "two.sided",
            method = paste(tools::toTitleCase(input$cor_method), "correlation (simple)"),
            data.name = paste(input$cor_var1, "and", input$cor_var2)
          )
          class(cor_test) <- "htest"
        })
      })
      
      if (is.null(cor_test)) {
        return(list(
          error = "Correlation calculation failed with all methods",
          success = FALSE
        ))
      }
      
      # 8. CREATE RESULTS TABLE
      # Calculate actual number of pairs used
      n_pairs <- sum(!is.na(var1_final) & !is.na(var2_final))
      
      results_df <- data.frame(
        Statistic = c(
          "Correlation Method",
          "Correlation Coefficient",
          paste(input$cor_conf * 100, "% CI Lower"),
          paste(input$cor_conf * 100, "% CI Upper"),
          "Test Statistic",
          "Degrees of Freedom",
          "P-value",
          "Interpretation",
          "Number of Complete Pairs",
          "Variable 1 Mean",
          "Variable 1 SD",
          "Variable 2 Mean",
          "Variable 2 SD",
          "Data Conversion Method (Var1)",
          "Data Conversion Method (Var2)",
          "Missing Value Handling"
        ),
        Value = c(
          tools::toTitleCase(input$cor_method),
          round(cor_test$estimate, 4),
          ifelse(!is.na(cor_test$conf.int[1]), round(cor_test$conf.int[1], 4), "N/A"),
          ifelse(!is.na(cor_test$conf.int[2]), round(cor_test$conf.int[2], 4), "N/A"),
          ifelse(!is.na(cor_test$statistic), round(cor_test$statistic, 4), "N/A"),
          ifelse(!is.na(cor_test$parameter), round(cor_test$parameter, 2), "N/A"),
          ifelse(!is.na(cor_test$p.value), 
                 ifelse(cor_test$p.value < 0.001, "<0.001", round(cor_test$p.value, 4)), 
                 "N/A"),
          ifelse(abs(cor_test$estimate) < 0.3, "Weak correlation",
                 ifelse(abs(cor_test$estimate) < 0.7, "Moderate correlation", 
                        "Strong correlation")),
          n_pairs,
          round(mean(var1_final, na.rm = TRUE), 3),
          round(sd(var1_final, na.rm = TRUE), 3),
          round(mean(var2_final, na.rm = TRUE), 3),
          round(sd(var2_final, na.rm = TRUE), 3),
          var1_conversion$method,
          var2_conversion$method,
          "Pairwise complete observations"
        ),
        stringsAsFactors = FALSE
      )
      
      # 9. CREATE INTERPRETATION
      r_value <- cor_test$estimate
      strength <- ifelse(abs(r_value) < 0.3, "weak",
                         ifelse(abs(r_value) < 0.7, "moderate", "strong"))
      direction <- ifelse(r_value > 0, "positive", "negative")
      
      interpretation <- list(
        strength = strength,
        direction = direction,
        r_value = r_value,
        p_value = cor_test$p.value,
        ci_lower = cor_test$conf.int[1],
        ci_upper = cor_test$conf.int[2],
        n_pairs = n_pairs,
        method = input$cor_method
      )
      
      # 10. RETURN COMPREHENSIVE RESULTS
      list(
        test = cor_test,
        data = data.frame(
          var1 = var1_final,
          var2 = var2_final,
          index = common_indices
        ),
        table = results_df,
        interpretation = interpretation,
        conversion_info = list(
          var1 = var1_conversion,
          var2 = var2_conversion
        ),
        success = TRUE
      )
      
    }, error = function(e) {
      # COMPREHENSIVE ERROR HANDLING
      error_details <- paste(
        "Correlation analysis failed with error: ", e$message,
        "\n\nTroubleshooting steps:",
        "\n1. Check that variables contain numeric data",
        "\n2. Ensure variables have at least 3 common observations after removing missing values",
        "\n3. Try recoding categorical variables to numeric",
        "\n4. Remove non-numeric characters from data",
        "\n5. Check for extreme outliers that might affect correlation",
        "\n6. Consider using Chi-square test for categorical variables"
      )
      
      return(list(
        error = error_details,
        success = FALSE
      ))
    })
  })
    # Add this function to check data types
  check_variable_types <- function(df, var1, var2) {
    type1 <- class(df[[var1]])[1]
    type2 <- class(df[[var2]])[1]
    unique1 <- length(unique(df[[var1]]))
    unique2 <- length(unique(df[[var2]]))
    
    cat("VARIABLE TYPE DIAGNOSTICS\n")
    cat("==========================\n")
    cat(var1, ":\n")
    cat("  Type:", type1, "\n")
    cat("  Unique values:", unique1, "\n")
    cat("  Sample values:", paste(head(df[[var1]], 5), collapse = ", "), "\n")
    cat(var2, ":\n")
    cat("  Type:", type2, "\n")
    cat("  Unique values:", unique2, "\n")
    cat("  Sample values:", paste(head(df[[var2]], 5), collapse = ", "), "\n")
  }
  # Correlation diagnostic information
  output$cor_diagnostic <- renderPrint({
    req(cor_results())
    
    results <- cor_results()
    
    if (!results$success) {
      cat("DIAGNOSTIC INFORMATION\n")
      cat("=====================\n")
      cat("Error:", results$error, "\n")
    } else if (!is.null(results$data_info)) {
      cat("DATA DIAGNOSTICS\n")
      cat("================\n")
      cat("Original variable types:", paste(results$data_info$original_types, collapse = ", "), "\n")
      cat("Converted variable types:", paste(results$data_info$converted_types, collapse = ", "), "\n")
      cat("Original N:", results$data_info$n_original, "\n")
      cat("Converted N:", results$data_info$n_converted, "\n")
      cat("Data loss:", results$data_info$n_original - results$data_info$n_converted, "observations\n")
      cat("\nFirst few values of converted data:\n")
      print(head(results$data))
    }
  })
  # Correlation Results Output
  output$cor_results <- renderPrint({
    req(cor_results())
    
    results <- cor_results()
    
    if (!results$success) {
      cat("CORRELATION ANALYSIS - ERROR\n")
      cat("============================\n")
      cat(results$error, "\n")
      return()
    }
    
    cat("CORRELATION ANALYSIS\n")
    cat("====================\n\n")
    
    cat("Variables:\n")
    cat("- X Variable:", input$cor_var1, "\n")
    cat("- Y Variable:", input$cor_var2, "\n")
    cat("- Method:", tools::toTitleCase(input$cor_method), "\n\n")
    
    cat("Descriptive Statistics:\n")
    cat("- N =", nrow(results$data), "\n")
    cat("-", input$cor_var1, "Mean =", round(mean(results$data$var1), 3), 
        ", SD =", round(sd(results$data$var1), 3), "\n")
    cat("-", input$cor_var2, "Mean =", round(mean(results$data$var2), 3), 
        ", SD =", round(sd(results$data$var2), 3), "\n\n")
    
    cat("Correlation Results:\n")
    print(results$test)
    
    cat("\nInterpretation:\n")
    cat("- Correlation coefficient (r) =", round(results$interpretation$r_value, 3), "\n")
    cat("- This indicates a", results$interpretation$strength, 
        results$interpretation$direction, "relationship.\n")
    if (results$interpretation$p_value < 0.05) {
      cat("- The correlation is statistically significant (p =", 
          round(results$interpretation$p_value, 4), ").\n")
    } else {
      cat("- The correlation is not statistically significant (p =", 
          round(results$interpretation$p_value, 4), ").\n")
    }
    cat("- We are", input$cor_conf * 100, "% confident that the true correlation lies between",
        round(results$interpretation$ci_lower, 3), "and",
        round(results$interpretation$ci_upper, 3), ".\n")
  })
  
  # Correlation Plot
  output$cor_plot <- renderPlotly({
    req(cor_results(), input$cor_scatter)
    
    results <- cor_results()
    
    if (!results$success) {
      p <- plot_ly() %>%
        add_annotations(
          text = "Error: Unable to generate scatter plot",
          x = 0.5,
          y = 0.5,
          xref = "paper",
          yref = "paper",
          showarrow = FALSE,
          font = list(size = 16, color = "red")
        ) %>%
        layout(
          title = "Scatter Plot - Error",
          xaxis = list(showgrid = FALSE, zeroline = FALSE, showticklabels = FALSE),
          yaxis = list(showgrid = FALSE, zeroline = FALSE, showticklabels = FALSE)
        )
      return(p)
    }
    
    # Create scatter plot
    p <- plot_ly(results$data,
                 x = ~var1,
                 y = ~var2,
                 type = 'scatter',
                 mode = 'markers',
                 marker = list(
                   color = input$cor_point_color,
                   size = input$cor_point_size,
                   opacity = input$cor_point_alpha,
                   line = list(color = 'rgba(0,0,0,0.2)', width = 1)
                 ),
                 hoverinfo = 'text',
                 text = ~paste(input$cor_var1, ":", round(var1, 2),
                               '<br>', input$cor_var2, ":", round(var2, 2)),
                 name = 'Data Points') %>%
      layout(
        title = paste("Scatter Plot:", input$cor_var1, "vs", input$cor_var2),
        xaxis = list(title = input$cor_var1),
        yaxis = list(title = input$cor_var2),
        hoverlabel = list(font = list(size = 12)),
        showlegend = TRUE
      )
    
    # Add trend line if requested
    if (input$cor_fit) {
      # Fit linear model
      fit <- lm(var2 ~ var1, data = results$data)
      x_range <- range(results$data$var1, na.rm = TRUE)
      x_seq <- seq(x_range[1], x_range[2], length.out = 100)
      pred_data <- data.frame(var1 = x_seq)
      pred_data$var2 <- predict(fit, newdata = pred_data)
      
      p <- p %>%
        add_trace(x = ~var1, y = ~var2,
                  data = pred_data,
                  type = 'scatter',
                  mode = 'lines',
                  line = list(color = input$cor_line_color, 
                              width = input$cor_line_size),
                  name = 'Trend Line',
                  hoverinfo = 'skip')
      
      # Add equation to plot
      slope <- round(coef(fit)[2], 3)
      intercept <- round(coef(fit)[1], 3)
      r_squared <- round(summary(fit)$r.squared, 3)
      
      p <- p %>%
        add_annotations(
          x = 0.05,
          y = 0.95,
          xref = "paper",
          yref = "paper",
          text = paste0("y = ", intercept, " + ", slope, "x<br>R² = ", r_squared),
          showarrow = FALSE,
          font = list(size = 12, color = input$cor_line_color),
          bgcolor = "rgba(255,255,255,0.8)",
          bordercolor = input$cor_line_color,
          borderwidth = 1,
          borderpad = 4
        )
    }
    
    # Add correlation coefficient annotation
    p <- p %>%
      add_annotations(
        x = 0.05,
        y = 0.85,
        xref = "paper",
        yref = "paper",
        text = paste0("r = ", round(results$interpretation$r_value, 3)),
        showarrow = FALSE,
        font = list(size = 14, color = "#0C5EA8", weight = "bold"),
        bgcolor = "rgba(255,255,255,0.8)",
        bordercolor = "#0C5EA8",
        borderwidth = 1,
        borderpad = 4
      )
    
    return(p)
  })
  
  # Correlation Table
  output$cor_table <- renderDT({
    req(cor_results())
    
    results <- cor_results()
    
    if (!results$success) {
      datatable(
        data.frame(Error = results$error),
        options = list(dom = 't'),
        rownames = FALSE,
        caption = "Correlation Analysis - Error"
      )
    } else {
      datatable(
        results$table,
        options = list(
          pageLength = 15,
          dom = 'Blfrtip',
          scrollX = TRUE,
          buttons = c('copy', 'csv', 'excel')
        ),
        rownames = FALSE,
        caption = "Correlation Analysis Results",
        extensions = 'Buttons'
      ) %>%
        formatStyle(
          'Value',
          backgroundColor = styleEqual(
            c("Weak correlation", "Moderate correlation", "Strong correlation"),
            c('#FFE5E5', '#FFF3CD', '#D4EDDA')
          )
        )
    }
  })
  
  # Download handler for correlation results (Word)
  output$download_cor <- downloadHandler(
    filename = function() {
      paste("correlation_analysis_", input$cor_var1, "_", input$cor_var2, "_", 
            Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(cor_results())
      
      results <- cor_results()
      
      if (!results$success) {
        # Create error document
        ft <- flextable(data.frame(Error = results$error)) %>%
          set_caption("Correlation Analysis - Error") %>%
          theme_zebra()
        
        doc <- read_docx() %>% 
          body_add_flextable(ft)
      } else {
        # Create comprehensive report
        doc <- read_docx()
        
        # Title and overview
        doc <- doc %>%
          body_add_par("CORRELATION ANALYSIS REPORT", style = "heading 1") %>%
          body_add_par(paste("Variable 1 (X):", input$cor_var1), style = "Normal") %>%
          body_add_par(paste("Variable 2 (Y):", input$cor_var2), style = "Normal") %>%
          body_add_par(paste("Method:", tools::toTitleCase(input$cor_method)), style = "Normal") %>%
          body_add_par(paste("Confidence Level:", input$cor_conf * 100, "%"), style = "Normal") %>%
          body_add_par("", style = "Normal")
        
        # Descriptive statistics
        doc <- doc %>%
          body_add_par("DESCRIPTIVE STATISTICS", style = "heading 2")
        
        desc_stats <- data.frame(
          Statistic = c("Number of Observations", 
                        paste(input$cor_var1, "Mean"),
                        paste(input$cor_var1, "SD"),
                        paste(input$cor_var2, "Mean"),
                        paste(input$cor_var2, "SD")),
          Value = c(
            nrow(results$data),
            round(mean(results$data$var1), 3),
            round(sd(results$data$var1), 3),
            round(mean(results$data$var2), 3),
            round(sd(results$data$var2), 3)
          )
        )
        
        ft_desc <- flextable(desc_stats) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_desc)
        
        # Correlation results
        doc <- doc %>%
          body_add_par("CORRELATION RESULTS", style = "heading 2")
        
        ft_cor <- flextable(results$table) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_cor)
        
        # Interpretation
        doc <- doc %>%
          body_add_par("INTERPRETATION", style = "heading 2")
        
        interpretation_text <- paste(
          "The correlation analysis revealed a ", 
          results$interpretation$strength, " ", 
          results$interpretation$direction, " relationship between ", 
          input$cor_var1, " and ", input$cor_var2, 
          " (r = ", round(results$interpretation$r_value, 3), 
          ifelse(results$interpretation$p_value < 0.05, 
                 ", p < 0.05, statistically significant).", 
                 ", p > 0.05, not statistically significant)."),
          " The ", input$cor_conf * 100, "% confidence interval for the correlation coefficient is [",
          round(results$interpretation$ci_lower, 3), ", ",
          round(results$interpretation$ci_upper, 3), "].",
          sep = ""
        )
        
        doc <- doc %>% 
          body_add_par(interpretation_text, style = "Normal") %>%
          body_add_par("", style = "Normal")
        
        # Statistical significance note
        if (results$interpretation$p_value < 0.05) {
          doc <- doc %>%
            body_add_par("Statistical Significance:", style = "heading 3") %>%
            body_add_par("The correlation is statistically significant (p < 0.05), suggesting that the observed relationship is unlikely to be due to chance.", style = "Normal")
        } else {
          doc <- doc %>%
            body_add_par("Statistical Significance:", style = "heading 3") %>%
            body_add_par("The correlation is not statistically significant (p > 0.05), suggesting that any observed relationship could be due to chance.", style = "Normal")
        }
        
        # Practical significance note
        doc <- doc %>%
          body_add_par("Practical Significance:", style = "heading 3")
        
        practical_text <- if (abs(results$interpretation$r_value) < 0.3) {
          "The correlation is weak, suggesting little practical relationship between the variables."
        } else if (abs(results$interpretation$r_value) < 0.7) {
          "The correlation is moderate, suggesting a meaningful but not strong relationship between the variables."
        } else {
          "The correlation is strong, suggesting a substantial relationship between the variables."
        }
        
        doc <- doc %>% body_add_par(practical_text, style = "Normal")
        
        # Assumptions check
        doc <- doc %>%
          body_add_par("ASSUMPTIONS CHECK", style = "heading 2") %>%
          body_add_par("For Pearson correlation, the key assumptions are:", style = "Normal") %>%
          body_add_par("1. Linearity: Relationship between variables should be linear", style = "Normal") %>%
          body_add_par("2. Homoscedasticity: Constant variance of residuals", style = "Normal") %>%
          body_add_par("3. Normality: Variables should be approximately normally distributed", style = "Normal") %>%
          body_add_par("Note: Kendall and Spearman correlations are non-parametric and do not require normality assumptions.", style = "Normal")
      }
      
      print(doc, target = file)
    }
  )
  
  # Download handler for correlation plot
  output$download_cor_plot <- downloadHandler(
    filename = function() {
      paste("correlation_plot_", input$cor_var1, "_", input$cor_var2, "_", 
            Sys.Date(), ".png", sep = "")
    },
    content = function(file) {
      req(cor_results(), input$cor_scatter)
      
      results <- cor_results()
      
      if (!results$success) {
        # Create error plot
        p <- ggplot() +
          annotate("text", x = 0.5, y = 0.5, 
                   label = "Error: Unable to generate correlation plot", 
                   size = 6, color = "red") +
          theme_void()
      } else {
        # Create high-quality scatter plot with ggplot2
        p <- ggplot(results$data, aes(x = var1, y = var2)) +
          geom_point(color = input$cor_point_color, 
                     size = input$cor_point_size,
                     alpha = input$cor_point_alpha) +
          labs(
            title = paste("Scatter Plot:", input$cor_var1, "vs", input$cor_var2),
            x = input$cor_var1,
            y = input$cor_var2,
            caption = paste("Correlation (", tools::toTitleCase(input$cor_method), 
                            "): r = ", round(results$interpretation$r_value, 3),
                            ", p = ", round(results$interpretation$p_value, 4))
          ) +
          theme_minimal() +
          theme(
            plot.title = element_text(size = 16, face = "bold", hjust = 0.5),
            axis.title = element_text(size = 14, face = "bold"),
            axis.text = element_text(size = 12),
            plot.caption = element_text(size = 11, face = "italic", hjust = 0.5),
            panel.grid.minor = element_blank(),
            panel.border = element_rect(color = "black", fill = NA, size = 0.5)
          )
        
        # Add trend line if requested
        if (input$cor_fit) {
          p <- p +
            geom_smooth(method = "lm", 
                        se = FALSE, 
                        color = input$cor_line_color,
                        size = input$cor_line_size)
          
          # Add equation and R-squared
          fit <- lm(var2 ~ var1, data = results$data)
          eq <- paste0("y = ", round(coef(fit)[1], 3), 
                       ifelse(coef(fit)[2] >= 0, " + ", " - "), 
                       abs(round(coef(fit)[2], 3)), "x")
          r2 <- paste0("R² = ", round(summary(fit)$r.squared, 3))
          
          p <- p +
            annotate("text", 
                     x = min(results$data$var1) + 0.1 * diff(range(results$data$var1)),
                     y = max(results$data$var2) - 0.1 * diff(range(results$data$var2)),
                     label = paste(eq, r2, sep = "\n"),
                     color = input$cor_line_color,
                     size = 4,
                     fontface = "bold")
        }
      }
      
      ggsave(file, plot = p, device = "png", width = 10, height = 8, dpi = 300, bg = "white")
    }
  )
  
  # Ultra-robust Central Tendency Analysis
  central_results <- eventReactive(input$run_central, {
    req(input$central_var, processed_data())
    
    tryCatch({
      df <- processed_data()
      
      # Double-check variable exists
      if (!input$central_var %in% names(df)) {
        return(list(
          table = data.frame(Error = "Variable not found in dataset"),
          success = FALSE
        ))
      }
      
      # Safely get the variable
      var_data <- df[[input$central_var]]
      
      # Handle non-numeric data
      if (!is.numeric(var_data)) {
        return(list(
          table = data.frame(Error = "Selected variable is not numeric"),
          success = FALSE
        ))
      }
      
      # Handle empty or all-NA data
      if (length(var_data) == 0 || all(is.na(var_data))) {
        return(list(
          table = data.frame(Error = "Variable contains no data"),
          success = FALSE
        ))
      }
      
      # Handle stratification variable if selected
      stratify_data <- NULL
      group_levels <- "Overall"
      
      if (input$central_stratify != "none" && input$central_stratify %in% names(df)) {
        stratify_data <- df[[input$central_stratify]]
        
        # Convert to factor and handle NAs for stratification
        if (all(is.na(stratify_data))) {
          return(list(
            table = data.frame(Error = "Stratification variable contains no valid data"),
            success = FALSE
          ))
        }
        
        stratify_data <- as.factor(stratify_data)
        stratify_data[is.na(stratify_data)] <- "Missing"
        group_levels <- levels(stratify_data)
      }
      
      # Remove NA values for calculations
      if (!is.null(stratify_data)) {
        clean_data <- var_data[!is.na(var_data) & !is.na(stratify_data)]
        clean_stratify <- stratify_data[!is.na(var_data) & !is.na(stratify_data)]
      } else {
        clean_data <- var_data[!is.na(var_data)]
        clean_stratify <- NULL
      }
      
      if (length(clean_data) == 0) {
        return(list(
          table = data.frame(Error = "No valid numeric data after removing missing values"),
          success = FALSE
        ))
      }
      
      # Calculate central tendency measures with stratification
      get_mode <- function(x) {
        if (length(x) == 0) return(NA)
        ux <- unique(x)
        ux[which.max(tabulate(match(x, ux)))]
      }
      
      calculate_stats <- function(data, group_name = "Overall") {
        data.frame(
          Group = group_name,
          Mean = mean(data, na.rm = TRUE),
          Median = median(data, na.rm = TRUE),
          SD = sd(data, na.rm = TRUE),
          Variance = var(data, na.rm = TRUE),
          Min = min(data, na.rm = TRUE),
          Max = max(data, na.rm = TRUE),
          Q1 = quantile(data, 0.25, na.rm = TRUE),
          Q3 = quantile(data, 0.75, na.rm = TRUE),
          IQR = IQR(data, na.rm = TRUE),
          Skewness = ifelse(length(data) > 2, moments::skewness(data, na.rm = TRUE), NA),
          Kurtosis = ifelse(length(data) > 3, moments::kurtosis(data, na.rm = TRUE), NA),
          Mode = get_mode(data),
          Count = length(data),
          Missing = sum(is.na(data)),
          stringsAsFactors = FALSE
        )
      }
      
      if (!is.null(clean_stratify)) {
        # Calculate statistics for each group
        central_table <- do.call(rbind, lapply(levels(clean_stratify), function(level) {
          group_data <- clean_data[clean_stratify == level]
          if (length(group_data) > 0) {
            calculate_stats(group_data, level)
          }
        }))
        
        # Add overall statistics
        overall_stats <- calculate_stats(clean_data, "Overall")
        central_table <- rbind(overall_stats, central_table)
        
      } else {
        # Calculate overall statistics only
        central_table <- calculate_stats(clean_data, "Overall")
      }
      
      # Format numeric values
      numeric_cols <- c("Mean", "Median", "SD", "Variance", "Min", "Max", "Q1", "Q3", "IQR", "Skewness", "Kurtosis")
      central_table[numeric_cols] <- lapply(central_table[numeric_cols], function(x) round(x, 4))
      
      list(
        table = central_table, 
        data = clean_data, 
        stratify = clean_stratify,
        group_levels = group_levels,
        success = TRUE
      )
      
    }, error = function(e) {
      return(list(
        table = data.frame(Error = paste("Error:", e$message)),
        success = FALSE
      ))
    })
  })
  
  output$central_table <- renderDT({
    req(central_results())
    
    results <- central_results()
    
    if (!results$success) {
      datatable(results$table, 
                options = list(dom = 't'),
                rownames = FALSE,
                caption = "Central Tendency Analysis - Error")
    } else {
      datatable(results$table,
                options = list(
                  pageLength = 20,
                  dom = 'Blfrtip',
                  scrollX = TRUE,
                  buttons = c('copy', 'csv', 'excel')
                ),
                rownames = FALSE,
                caption = "Central Tendency Measures",
                extensions = 'Buttons'
      ) %>%
        formatRound(columns = c('Mean', 'Median', 'SD', 'Variance', 'Min', 'Max', 
                                'Q1', 'Q3', 'IQR', 'Skewness', 'Kurtosis'), 
                    digits = 4) %>%
        formatStyle(
          'Mean',
          background = styleColorBar(range(results$table$Mean, na.rm = TRUE), '#0C5EA8'),
          backgroundSize = '100% 90%',
          backgroundRepeat = 'no-repeat',
          backgroundPosition = 'center'
        )
    }
  })
  
  output$central_boxplot <- renderPlotly({
    req(central_results())
    
    results <- central_results()
    
    if (!results$success) {
      p <- plot_ly() %>%
        add_annotations(
          text = results$table$Error[1],
          x = 0.5,
          y = 0.5,
          xref = "paper",
          yref = "paper",
          showarrow = FALSE,
          font = list(size = 16, color = "red")
        ) %>%
        layout(
          title = "Boxplot - Error",
          xaxis = list(showgrid = FALSE, zeroline = FALSE, showticklabels = FALSE),
          yaxis = list(showgrid = FALSE, zeroline = FALSE, showticklabels = FALSE)
        )
      return(p)
    }
    
    if (!is.null(results$stratify)) {
      # Create grouped box plot
      plot_data <- data.frame(
        value = results$data,
        group = results$stratify
      )
      
      p <- plot_ly(plot_data, 
                   x = ~group, 
                   y = ~value, 
                   type = "box",
                   color = ~group,
                   boxpoints = "all",
                   jitter = 0.3,
                   pointpos = -1.8,
                   marker = list(size = 3, opacity = 0.7),
                   line = list(width = 1.5)) %>%
        layout(
          title = list(text = paste(input$central_title, "- Boxplot by", input$central_stratify), 
                       font = list(size = input$central_title_size)),
          yaxis = list(title = input$central_var,
                       titlefont = list(size = input$central_y_size)),
          xaxis = list(title = input$central_stratify,
                       titlefont = list(size = input$central_x_size)),
          margin = list(b = 80, l = 60, r = 40, t = 60),
          showlegend = FALSE
        )
    } else {
      # Create single box plot
      p <- plot_ly(y = ~results$data,
                   type = "box",
                   name = input$central_var,
                   boxpoints = "all",
                   jitter = 0.3,
                   pointpos = -1.8,
                   marker = list(color = input$central_color, 
                                 size = 3,
                                 opacity = 0.7),
                   line = list(color = input$central_color, width = 1.5),
                   fillcolor = paste0(input$central_color, "33")) %>%
        layout(
          title = list(text = paste(input$central_title, "- Boxplot"), 
                       font = list(size = input$central_title_size)),
          yaxis = list(title = input$central_var,
                       titlefont = list(size = input$central_y_size)),
          xaxis = list(title = "", showticklabels = FALSE),
          margin = list(l = 60, r = 40, t = 60),
          showlegend = FALSE
        )
    }
    
    return(p)
  })
  
  output$central_histogram <- renderPlotly({
    req(central_results())
    
    results <- central_results()
    
    if (!results$success) {
      p <- plot_ly() %>%
        add_annotations(
          text = results$table$Error[1],
          x = 0.5,
          y = 0.5,
          xref = "paper",
          yref = "paper",
          showarrow = FALSE,
          font = list(size = 16, color = "red")
        ) %>%
        layout(
          title = "Histogram - Error",
          xaxis = list(showgrid = FALSE, zeroline = FALSE, showticklabels = FALSE),
          yaxis = list(showgrid = FALSE, zeroline = FALSE, showticklabels = FALSE)
        )
      return(p)
    }
    
    if (!is.null(results$stratify)) {
      # Create grouped histogram
      plot_data <- data.frame(
        value = results$data,
        group = results$stratify
      )
      
      p <- plot_ly(plot_data) %>%
        add_histogram(x = ~value, color = ~group, opacity = 0.7,
                      marker = list(line = list(color = 'white', width = 1))) %>%
        layout(
          title = list(text = paste(input$central_title, "- Histogram by", input$central_stratify), 
                       font = list(size = input$central_title_size)),
          xaxis = list(title = input$central_var,
                       titlefont = list(size = input$central_x_size)),
          yaxis = list(title = "Frequency",
                       titlefont = list(size = input$central_y_size)),
          barmode = "overlay",
          bargap = 0.1,
          margin = list(b = 80, l = 60, r = 40, t = 60),
          legend = list(orientation = "h", y = -0.2)
        )
    } else {
      # Create single histogram
      p <- plot_ly(x = ~results$data,
                   type = "histogram",
                   nbinsx = input$central_bins,
                   marker = list(color = input$central_color,
                                 line = list(color = "white", width = 1)),
                   opacity = 0.7) %>%
        layout(
          title = list(text = paste(input$central_title, "- Histogram"), 
                       font = list(size = input$central_title_size)),
          xaxis = list(title = input$central_var,
                       titlefont = list(size = input$central_x_size)),
          yaxis = list(title = "Frequency",
                       titlefont = list(size = input$central_y_size)),
          bargap = 0.1,
          margin = list(b = 80, l = 60, r = 40, t = 60)
        )
    }
    
    return(p)
  })
  
  # CORRECTED DOWNLOAD HANDLERS FOR DESCRIPTIVE ANALYSIS
  
  # Frequency Analysis Download
  output$download_freq <- downloadHandler(
    filename = function() {
      if(input$freq_type == "Table") {
        paste("frequency_analysis_", input$freq_var, "_", Sys.Date(), ".docx", sep = "")
      } else {
        paste("frequency_plot_", input$freq_var, "_", Sys.Date(), ".png", sep = "")
      }
    },
    content = function(file) {
      req(freq_results())
      
      results <- freq_results()
      
      if(input$freq_type == "Table") {
        if (!results$success) {
          ft <- flextable(data.frame(Error = "Unable to generate frequency table due to error")) %>%
            set_caption("Frequency Analysis - Error") %>%
            theme_zebra()
          
          doc <- read_docx() %>% 
            body_add_flextable(ft)
        } else {
          ft <- flextable(results$table) %>%
            set_caption(paste("Frequency Distribution -", input$freq_var)) %>%
            theme_zebra() %>%
            autofit()
          
          doc <- read_docx() %>% 
            body_add_flextable(ft)
        }
        print(doc, target = file)
        
      } else {
        if (!results$success) {
          p <- ggplot() +
            annotate("text", x = 1, y = 1, 
                     label = "Error: Unable to generate frequency plot", 
                     size = 6, color = "red") +
            theme_void()
        } else {
          p <- ggplot(results$table, aes(x = Category, y = Frequency)) +
            geom_bar(stat = "identity", fill = input$freq_color, alpha = 0.8) +
            geom_text(aes(label = Frequency), vjust = -0.5, size = 4) +
            labs(title = input$freq_title,
                 x = input$freq_var,
                 y = "Frequency") +
            theme_minimal() +
            theme(
              plot.title = element_text(size = input$freq_title_size, face = "bold", hjust = 0.5),
              axis.title.x = element_text(size = input$freq_x_size, face = "bold"),
              axis.title.y = element_text(size = input$freq_y_size, face = "bold"),
              axis.text.x = element_text(angle = 45, hjust = 1, size = 11),
              axis.text.y = element_text(size = 11),
              panel.grid.major = element_line(color = "grey90"),
              panel.grid.minor = element_blank()
            )
        }
        ggsave(file, plot = p, device = "png", width = 10, height = 6, dpi = 300)
      }
    }
  )
  
  # Central Tendency Download
  output$download_central <- downloadHandler(
    filename = function() {
      if(input$central_type == "Table") {
        paste("central_tendency_", input$central_var, "_", Sys.Date(), ".docx", sep = "")
      } else if(input$central_type == "Boxplot") {
        paste("boxplot_", input$central_var, "_", Sys.Date(), ".png", sep = "")
      } else {
        paste("histogram_", input$central_var, "_", Sys.Date(), ".png", sep = "")
      }
    },
    content = function(file) {
      req(central_results())
      
      results <- central_results()
      
      if(input$central_type == "Table") {
        if (!results$success) {
          ft <- flextable(data.frame(Error = "Unable to generate central tendency table due to error")) %>%
            set_caption("Central Tendency Analysis - Error") %>%
            theme_zebra()
          
          doc <- read_docx() %>% 
            body_add_flextable(ft)
        } else {
          ft <- flextable(results$table) %>%
            set_caption(paste("Central Tendency Measures -", input$central_var)) %>%
            theme_zebra() %>%
            autofit()
          
          doc <- read_docx() %>% 
            body_add_flextable(ft)
        }
        print(doc, target = file)
        
      } else if(input$central_type == "Boxplot") {
        if (!results$success) {
          p <- ggplot() +
            annotate("text", x = 1, y = 1, 
                     label = "Error: Unable to generate boxplot", 
                     size = 6, color = "red") +
            theme_void()
        } else {
          df_plot <- data.frame(value = results$data)
          if (!is.null(results$stratify)) {
            df_plot$group <- results$stratify
            p <- ggplot(df_plot, aes(x = group, y = value, fill = group)) +
              geom_boxplot(alpha = 0.7, outlier.color = "red", outlier.size = 2) +
              stat_summary(fun = mean, geom = "point", shape = 18, size = 3, color = "blue") +
              labs(title = paste(input$central_title, "- Boxplot by", input$central_stratify),
                   y = input$central_var,
                   x = input$central_stratify) +
              theme_minimal() +
              theme(
                plot.title = element_text(size = input$central_title_size, face = "bold", hjust = 0.5),
                axis.title.x = element_text(size = 14, face = "bold"),
                axis.title.y = element_text(size = 14, face = "bold"),
                legend.position = "none"
              ) +
              scale_fill_brewer(palette = "Set2")
          } else {
            p <- ggplot(df_plot, aes(y = value)) +
              geom_boxplot(fill = input$central_color, alpha = 0.7, outlier.color = "red", outlier.size = 2) +
              stat_summary(fun = mean, geom = "point", shape = 18, size = 3, color = "blue") +
              labs(title = paste(input$central_title, "- Boxplot"),
                   y = input$central_var) +
              theme_minimal() +
              theme(
                plot.title = element_text(size = input$central_title_size, face = "bold", hjust = 0.5),
                axis.title.y = element_text(size = 14, face = "bold"),
                axis.title.x = element_blank(),
                axis.text.x = element_blank(),
                axis.ticks.x = element_blank()
              )
          }
        }
        ggsave(file, plot = p, device = "png", width = 8, height = 6, dpi = 300)
        
      } else {
        if (!results$success) {
          p <- ggplot() +
            annotate("text", x = 1, y = 1, 
                     label = "Error: Unable to generate histogram", 
                     size = 6, color = "red") +
            theme_void()
        } else {
          df_plot <- data.frame(value = results$data)
          if (!is.null(results$stratify)) {
            df_plot$group <- results$stratify
            p <- ggplot(df_plot, aes(x = value, fill = group)) +
              geom_histogram(alpha = 0.7, position = "identity", bins = input$central_bins) +
              labs(title = paste(input$central_title, "- Histogram by", input$central_stratify),
                   x = input$central_var, 
                   y = "Frequency",
                   fill = input$central_stratify) +
              theme_minimal() +
              theme(
                plot.title = element_text(size = input$central_title_size, face = "bold", hjust = 0.5),
                axis.title = element_text(size = 14, face = "bold"),
                legend.position = "right"
              ) +
              scale_fill_brewer(palette = "Set2")
          } else {
            p <- ggplot(df_plot, aes(x = value)) +
              geom_histogram(fill = input$central_color, bins = input$central_bins, 
                             alpha = 0.7, color = "white") +
              labs(title = paste(input$central_title, "- Histogram"),
                   x = input$central_var, 
                   y = "Frequency") +
              theme_minimal() +
              theme(
                plot.title = element_text(size = input$central_title_size, face = "bold", hjust = 0.5),
                axis.title = element_text(size = 14, face = "bold")
              )
          }
        }
        ggsave(file, plot = p, device = "png", width = 10, height = 6, dpi = 300)
      }
    }
  )
  
  # Ultra-robust Epidemic Curve Analysis
  epi_results <- eventReactive(input$run_epi, {
    req(input$epi_date, input$epi_case, processed_data())
    
    tryCatch({
      df <- processed_data()
      
      # Double-check variables exist
      if (!input$epi_date %in% names(df)) {
        return(list(
          data = data.frame(),
          error = "Date variable not found in dataset",
          success = FALSE
        ))
      }
      
      if (!input$epi_case %in% names(df)) {
        return(list(
          data = data.frame(),
          error = "Case count variable not found in dataset",
          success = FALSE
        ))
      }
      
      if (input$epi_group != "none" && !input$epi_group %in% names(df)) {
        return(list(
          data = data.frame(),
          error = "Grouping variable not found in dataset",
          success = FALSE
        ))
      }
      
      # Safely get the variables
      date_data <- df[[input$epi_date]]
      case_data <- df[[input$epi_case]]
      
      # Handle empty or all-NA data
      if (length(date_data) == 0 || all(is.na(date_data))) {
        return(list(
          data = data.frame(),
          error = "Date variable contains no data",
          success = FALSE
        ))
      }
      
      if (length(case_data) == 0 || all(is.na(case_data))) {
        return(list(
          data = data.frame(),
          error = "Case count variable contains no data",
          success = FALSE
        ))
      }
      
      # Convert date variable to character/factor for categorical handling
      df$date_converted <- as.character(date_data)
      
      # Handle missing values in date
      df$date_converted[is.na(df$date_converted)] <- "Missing"
      
      # Ensure case count is numeric
      df$case_numeric <- as.numeric(case_data)
      
      if (all(is.na(df$case_numeric))) {
        return(list(
          data = data.frame(),
          error = "Case count variable could not be converted to numeric",
          success = FALSE
        ))
      }
      
      # Handle grouping variable
      if (input$epi_group != "none") {
        group_data <- df[[input$epi_group]]
        if (all(is.na(group_data))) {
          return(list(
            data = data.frame(),
            error = "Grouping variable contains no valid data",
            success = FALSE
          ))
        }
        df$group_var <- as.character(group_data)
        df$group_var[is.na(df$group_var)] <- "Missing"
      }
      
      # Use date_converted directly as time_group for categorical aggregation
      df$time_group <- df$date_converted
      
      # For date-like categorical variables, try to convert to proper dates
      if (input$epi_interval %in% c("Day", "Week", "Month", "Year")) {
        tryCatch({
          sample_dates <- head(na.omit(unique(df$date_converted)), 10)
          
          looks_like_date <- any(
            grepl("\\d{1,4}[/-]\\d{1,2}[/-]\\d{1,4}", sample_dates) |
              grepl("\\d{1,2}[/-]\\d{1,2}[/-]\\d{2,4}", sample_dates) |
              grepl("jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec", tolower(sample_dates), ignore.case = TRUE)
          )
          
          if (looks_like_date) {
            convert_date <- function(x) {
              formats <- c("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%Y.%m.%d", 
                           "%d-%m-%Y", "%m-%d-%Y", "%B %d, %Y", "%d %B %Y",
                           "%b %d, %Y", "%d %b %Y", "%Y/%m/%d", "%m/%d/%y")
              
              for (fmt in formats) {
                parsed <- as.Date(x, format = fmt)
                if (!all(is.na(parsed))) {
                  return(parsed)
                }
              }
              return(as.Date(NA))
            }
            
            df$date_attempt <- convert_date(df$date_converted)
            
            if (!all(is.na(df$date_attempt))) {
              if (input$epi_interval == "Day") {
                df$time_group <- df$date_attempt
              } else if (input$epi_interval == "Week") {
                df$time_group <- lubridate::floor_date(df$date_attempt, "week")
              } else if (input$epi_interval == "Month") {
                df$time_group <- lubridate::floor_date(df$date_attempt, "month")
              } else {
                df$time_group <- lubridate::floor_date(df$date_attempt, "year")
              }
              
              df$time_group <- as.character(df$time_group)
            }
          }
        }, error = function(e) {
          message("Date conversion failed: ", e$message)
        })
      }
      
      # Aggregate data
      if (input$epi_group != "none") {
        epi_data <- df %>%
          group_by(time_group, group_var) %>%
          summarise(
            cases = sum(case_numeric, na.rm = TRUE),
            n_records = n(),
            .groups = "drop"
          ) %>%
          rename(Group = group_var)
      } else {
        epi_data <- df %>%
          group_by(time_group) %>%
          summarise(
            cases = sum(case_numeric, na.rm = TRUE),
            n_records = n(),
            .groups = "drop"
          )
      }
      
      # Remove any time groups with no cases
      epi_data <- epi_data %>%
        filter(cases > 0)
      
      if (nrow(epi_data) == 0) {
        return(list(
          data = data.frame(),
          error = "No cases found in the specified date range",
          success = FALSE
        ))
      }
      
      # Sort by time_group for proper ordering
      epi_data <- epi_data %>%
        arrange(time_group)
      
      # Calculate summary statistics
      summary_stats <- list(
        total_cases = sum(epi_data$cases),
        time_range = paste(min(epi_data$time_group), "to", max(epi_data$time_group)),
        n_time_periods = nrow(epi_data),
        peak_cases = max(epi_data$cases),
        peak_time = epi_data$time_group[which.max(epi_data$cases)]
      )
      
      list(
        data = epi_data,
        summary = summary_stats,
        success = TRUE
      )
      
    }, error = function(e) {
      return(list(
        data = data.frame(),
        error = paste("Error:", e$message),
        success = FALSE
      ))
    })
  })
  
  output$epi_curve <- renderPlotly({
    req(epi_results())
    
    results <- epi_results()
    
    if (!results$success) {
      p <- plot_ly() %>%
        add_annotations(
          text = results$error,
          x = 0.5,
          y = 0.5,
          xref = "paper",
          yref = "paper",
          showarrow = FALSE,
          font = list(size = 14, color = "red")
        ) %>%
        layout(
          title = "Epidemic Curve - Error",
          xaxis = list(showgrid = FALSE, zeroline = FALSE, showticklabels = FALSE),
          yaxis = list(showgrid = FALSE, zeroline = FALSE, showticklabels = FALSE)
        )
      return(p)
    }
    
    epi_data <- results$data
    
    # Create epidemic curve
    if (input$epi_group != "none") {
      p <- plot_ly(epi_data, 
                   x = ~time_group, 
                   y = ~cases, 
                   color = ~Group,
                   type = 'bar',
                   hoverinfo = 'text',
                   text = ~paste('Date:', time_group,
                                 '<br>Group:', Group,
                                 '<br>Cases:', cases,
                                 '<br>Records:', n_records)) %>%
        layout(
          title = list(text = input$epi_title, 
                       font = list(size = input$epi_title_size)),
          xaxis = list(title = "Date", tickangle = 45,
                       titlefont = list(size = input$epi_x_size)),
          yaxis = list(title = "Number of Cases",
                       titlefont = list(size = input$epi_y_size)),
          barmode = 'stack',
          margin = list(b = 100, l = 60, r = 40, t = 60),
          legend = list(orientation = "h", y = -0.2),
          hoverlabel = list(font = list(size = 12))
        )
    } else {
      p <- plot_ly(epi_data, 
                   x = ~time_group, 
                   y = ~cases,
                   type = 'bar',
                   marker = list(color = input$epi_color,
                                 line = list(color = 'rgba(0,0,0,0.2)', width = 1)),
                   hoverinfo = 'text',
                   text = ~paste('Date:', time_group,
                                 '<br>Cases:', cases,
                                 '<br>Records:', n_records)) %>%
        layout(
          title = list(text = input$epi_title, 
                       font = list(size = input$epi_title_size)),
          xaxis = list(title = "Date", tickangle = 45,
                       titlefont = list(size = input$epi_x_size)),
          yaxis = list(title = "Number of Cases",
                       titlefont = list(size = input$epi_y_size)),
          margin = list(b = 100, l = 60, r = 40, t = 60),
          showlegend = FALSE,
          hoverlabel = list(font = list(size = 12))
        )
    }
    
    return(p)
  })
  
  output$epi_summary <- renderUI({
    req(epi_results())
    
    results <- epi_results()
    
    if (!results$success) {
      return(
        div(
          class = "alert alert-danger",
          h4("Analysis Error"),
          p(results$error)
        )
      )
    }
    
    summary <- results$summary
    
    div(
      class = "card",
      div(
        class = "card-header",
        h4("Epidemic Summary Statistics", class = "card-title", style = "margin: 0;")
      ),
      div(
        class = "card-body",
        fluidRow(
          column(6,
                 tags$ul(
                   class = "list-unstyled",
                   tags$li(tags$strong("Total Cases:"), summary$total_cases),
                   tags$li(tags$strong("Date Range:"), summary$time_range),
                   tags$li(tags$strong("Time Periods:"), summary$n_time_periods)
                 )
          ),
          column(6,
                 tags$ul(
                   class = "list-unstyled",
                   tags$li(tags$strong("Peak Cases:"), summary$peak_cases),
                   tags$li(tags$strong("Peak Time:"), summary$peak_time)
                 )
          )
        )
      )
    )
  })
  
  output$download_epicurve <- downloadHandler(
    filename = function() {
      paste("epidemic_curve_", input$epi_interval, "_", Sys.Date(), ".png", sep = "")
    },
    content = function(file) {
      req(epi_results())
      
      results <- epi_results()
      
      if (!results$success) {
        p <- ggplot() +
          annotate("text", x = 1, y = 1, 
                   label = paste("Error:", results$error), 
                   size = 5, color = "red") +
          theme_void()
      } else {
        epi_data <- results$data
        
        if (input$epi_group != "none") {
          p <- ggplot(epi_data, aes(x = time_group, y = cases, fill = Group)) +
            geom_col(position = "stack") +
            scale_fill_brewer(palette = "Set1") +
            labs(
              title = input$epi_title,
              subtitle = paste("Time Interval:", input$epi_interval),
              x = "Date",
              y = "Number of Cases",
              fill = input$epi_group
            )
        } else {
          p <- ggplot(epi_data, aes(x = time_group, y = cases)) +
            geom_col(fill = input$epi_color, alpha = 0.8) +
            labs(
              title = input$epi_title,
              subtitle = paste("Time Interval:", input$epi_interval),
              x = "Date",
              y = "Number of Cases"
            )
        }
        
        p <- p +
          theme_minimal() +
          theme(
            plot.title = element_text(size = input$epi_title_size, face = "bold", hjust = 0.5),
            plot.subtitle = element_text(size = input$epi_title_size - 2, color = "gray50", hjust = 0.5),
            axis.title.x = element_text(size = input$epi_x_size, face = "bold"),
            axis.title.y = element_text(size = input$epi_y_size, face = "bold"),
            axis.text.x = element_text(angle = 45, hjust = 1, size = 10),
            axis.text.y = element_text(size = 10),
            legend.position = ifelse(input$epi_group != "none", "right", "none"),
            panel.grid.major = element_line(color = "grey90"),
            panel.grid.minor = element_blank(),
            panel.border = element_rect(color = "black", fill = NA, size = 0.5)
          )
      }
      
      ggsave(file, plot = p, device = "png", width = 12, height = 8, dpi = 300)
    }
  )
  
  output$download_epi_data <- downloadHandler(
    filename = function() {
      paste("epidemic_data_", input$epi_interval, "_", Sys.Date(), ".csv", sep = "")
    },
    content = function(file) {
      req(epi_results())
      
      results <- epi_results()
      
      if (!results$success) {
        write.csv(data.frame(Error = results$error), file, row.names = FALSE)
      } else {
        write.csv(results$data, file, row.names = FALSE)
      }
    }
  )
  
  # Incidence/Prevalence with action button
  ip_results <- eventReactive(input$run_ip, {
    req(input$ip_case, input$ip_pop, processed_data())
    df <- processed_data()
    
    if(input$ip_measure == "Incidence Rate") {
      req(input$ip_time)
      total_cases <- sum(df[[input$ip_case]], na.rm = TRUE)
      total_pop <- sum(df[[input$ip_pop]], na.rm = TRUE)
      total_time <- sum(df[[input$ip_time]], na.rm = TRUE)
      
      rate <- (total_cases / total_time) * 1000  # per 1000 person-time
      
      result <- data.frame(
        Measure = "Incidence Rate",
        Cases = total_cases,
        Population = total_pop,
        PersonTime = total_time,
        Rate = round(rate, 2),
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
        Rate = round(ci, 2),
        Units = "per 1000 population"
      )
    } else {
      total_cases <- sum(df[[input$ip_case]], na.rm = TRUE)
      total_pop <- sum(df[[input$ip_pop]], na.rm = TRUE)
      
      prevalence <- (total_cases / total_pop) * 100  # percentage
      
      result <- data.frame(
        Measure = "Prevalence",
        Cases = total_cases,
        Population = total_pop,
        Rate = round(prevalence, 2),
        Units = "percent"
      )
    }
    
    list(table = result)
  })
  
  output$ip_table <- renderDT({
    req(ip_results())
    datatable(ip_results()$table, 
              options = list(dom = 't'),
              rownames = FALSE,
              caption = "Disease Frequency Measures") %>%
      formatStyle(
        'Rate',
        backgroundColor = styleInterval(c(0.1, 1, 10), 
                                        c('#FFE5E5', '#FFF3CD', '#D4EDDA', '#D1ECF1'))
      )
  })
  
  output$download_ip <- downloadHandler(
    filename = "incidence_prevalence.docx",
    content = function(file) {
      ft <- flextable(ip_results()$table) %>%
        set_caption("Disease Frequency Measures") %>%
        theme_zebra() %>%
        autofit()
      
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # 4.0 COMPLETE DOWNLOAD HANDLERS
  
  # One-Way ANOVA Download
  output$download_anova <- downloadHandler(
    filename = function() {
      paste("oneway_anova_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(anova_results())
      results <- anova_results()
      
      doc <- read_docx()
      
      if (!results$success) {
        doc <- doc %>%
          body_add_par("ONE-WAY ANOVA ANALYSIS - ERROR", style = "heading 1") %>%
          body_add_par(results$error, style = "Normal")
      } else {
        doc <- doc %>%
          body_add_par("ONE-WAY ANOVA ANALYSIS REPORT", style = "heading 1") %>%
          body_add_par(paste("Outcome Variable:", input$anova_outcome), style = "Normal") %>%
          body_add_par(paste("Grouping Variable:", input$anova_group), style = "Normal") %>%
          body_add_par(paste("Number of Groups:", length(unique(processed_data()[[input$anova_group]]))), style = "Normal") %>%
          body_add_par("", style = "Normal")
        
        doc <- doc %>%
          body_add_par("ANOVA RESULTS", style = "heading 2")
        
        anova_df <- as.data.frame(results$summary[[1]])
        anova_df <- cbind(Source = rownames(anova_df), anova_df)
        rownames(anova_df) <- NULL
        
        ft_anova <- flextable(anova_df) %>%
          theme_zebra() %>%
          set_caption("One-Way ANOVA Table") %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_anova)
        
        doc <- doc %>%
          body_add_par("GROUP DESCRIPTIVES", style = "heading 2")
        
        df <- processed_data()
        descriptives <- df %>%
          group_by(Group = !!sym(input$anova_group)) %>%
          summarise(
            N = sum(!is.na(!!sym(input$anova_outcome))),
            Mean = round(mean(!!sym(input$anova_outcome), na.rm = TRUE), 3),
            SD = round(sd(!!sym(input$anova_outcome), na.rm = TRUE), 3),
            SE = round(sd(!!sym(input$anova_outcome), na.rm = TRUE) / sqrt(sum(!is.na(!!sym(input$anova_outcome)))), 3),
            .groups = "drop"
          )
        
        ft_desc <- flextable(descriptives) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_desc)
        
        if (!is.null(results$posthoc)) {
          doc <- doc %>%
            body_add_par("POST-HOC COMPARISONS", style = "heading 2")
          
          if (input$anova_posthoc_method == "tukey") {
            posthoc_df <- as.data.frame(results$posthoc$group_clean)
            posthoc_df <- cbind(Comparison = rownames(posthoc_df), posthoc_df)
            rownames(posthoc_df) <- NULL
            
            ft_posthoc <- flextable(posthoc_df) %>%
              theme_zebra() %>%
              set_caption("Tukey HSD Post-hoc Comparisons") %>%
              autofit()
            
            doc <- doc %>% body_add_flextable(ft_posthoc)
          } else {
            posthoc_text <- paste("Post-hoc method:", input$anova_posthoc_method)
            doc <- doc %>% body_add_par(posthoc_text, style = "Normal")
          }
        }
        
        if (!is.null(results$assumptions)) {
          doc <- doc %>%
            body_add_par("ASSUMPTIONS CHECK", style = "heading 2")
          
          assumptions_df <- data.frame(
            Test = c("Normality (Shapiro-Wilk)", "Homogeneity of Variance (Levene)"),
            Statistic = c(
              round(results$assumptions$normality$statistic, 4),
              round(results$assumptions$homogeneity$`F value`[1], 4)
            ),
            P_Value = c(
              round(results$assumptions$normality$p.value, 4),
              round(results$assumptions$homogeneity$`Pr(>F)`[1], 4)
            ),
            Interpretation = c(
              ifelse(results$assumptions$normality$p.value > 0.05, "Assumption met", "Assumption violated"),
              ifelse(results$assumptions$homogeneity$`Pr(>F)`[1] > 0.05, "Assumption met", "Assumption violated")
            )
          )
          
          ft_assumptions <- flextable(assumptions_df) %>%
            theme_zebra() %>%
            autofit()
          
          doc <- doc %>% body_add_flextable(ft_assumptions)
        }
        
        doc <- doc %>%
          body_add_par("INTERPRETATION", style = "heading 2")
        
        p_value <- results$summary[[1]]$`Pr(>F)`[1]
        f_value <- results$summary[[1]]$`F value`[1]
        
        interpretation <- if (p_value < 0.05) {
          paste("The one-way ANOVA revealed a statistically significant difference between groups ",
                "(F =", round(f_value, 3), ", p =", round(p_value, 4), "). ",
                "This suggests that there are significant differences in", input$anova_outcome, 
                "across the levels of", input$anova_group, ".")
        } else {
          paste("The one-way ANOVA did not reveal a statistically significant difference between groups ",
                "(F =", round(f_value, 3), ", p =", round(p_value, 4), "). ",
                "This suggests that there are no significant differences in", input$anova_outcome, 
                "across the levels of", input$anova_group, ".")
        }
        
        doc <- doc %>% body_add_par(interpretation, style = "Normal")
      }
      
      print(doc, target = file)
    }
  )
  
  # Two-Way ANOVA Download
  output$download_anova_2way <- downloadHandler(
    filename = function() {
      paste("twoway_anova_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(anova_2way_results())
      results <- anova_2way_results()
      
      doc <- read_docx()
      
      if (!results$success) {
        doc <- doc %>%
          body_add_par("TWO-WAY ANOVA ANALYSIS - ERROR", style = "heading 1") %>%
          body_add_par(results$error, style = "Normal")
      } else {
        doc <- doc %>%
          body_add_par("TWO-WAY ANOVA ANALYSIS REPORT", style = "heading 1") %>%
          body_add_par(paste("Outcome Variable:", input$anova_outcome_2way), style = "Normal") %>%
          body_add_par(paste("Factor 1:", input$anova_factor1), style = "Normal") %>%
          body_add_par(paste("Factor 2:", input$anova_factor2), style = "Normal") %>%
          body_add_par(paste("Interaction Term:", ifelse(input$anova_interaction, "Included", "Excluded")), style = "Normal") %>%
          body_add_par("", style = "Normal")
        
        doc <- doc %>%
          body_add_par("TWO-WAY ANOVA RESULTS", style = "heading 2")
        
        anova_df <- as.data.frame(results$summary[[1]])
        anova_df <- cbind(Source = rownames(anova_df), anova_df)
        rownames(anova_df) <- NULL
        
        ft_anova <- flextable(anova_df) %>%
          theme_zebra() %>%
          set_caption("Two-Way ANOVA Table") %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_anova)
        
        if (!is.null(results$simple_effects)) {
          doc <- doc %>%
            body_add_par("SIMPLE EFFECTS ANALYSIS", style = "heading 2")
          
          simple_effects_text <- capture.output(print(results$simple_effects))
          for (line in simple_effects_text) {
            doc <- doc %>% body_add_par(line, style = "Normal")
          }
        }
        
        doc <- doc %>%
          body_add_par("INTERPRETATION", style = "heading 2")
        
        p_values <- results$summary[[1]]$`Pr(>F)`
        sources <- rownames(results$summary[[1]])
        
        interpretation <- "Key findings:\n"
        
        for (i in 1:length(sources)) {
          if (!is.na(p_values[i])) {
            significance <- ifelse(p_values[i] < 0.05, "significant", "not significant")
            interpretation <- paste0(interpretation, 
                                     "- ", sources[i], ": ", significance, " (p = ", round(p_values[i], 4), ")\n")
          }
        }
        
        doc <- doc %>% body_add_par(interpretation, style = "Normal")
      }
      
      print(doc, target = file)
    }
  )
  
  # Welch ANOVA Download
  output$download_welch_anova <- downloadHandler(
    filename = function() {
      paste("welch_anova_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(welch_anova_results())
      results <- welch_anova_results()
      
      doc <- read_docx()
      
      if (!results$success) {
        doc <- doc %>%
          body_add_par("WELCH ANOVA ANALYSIS - ERROR", style = "heading 1") %>%
          body_add_par(results$error, style = "Normal")
      } else {
        doc <- doc %>%
          body_add_par("WELCH ANOVA ANALYSIS REPORT", style = "heading 1") %>%
          body_add_par(paste("Outcome Variable:", input$welch_outcome), style = "Normal") %>%
          body_add_par(paste("Grouping Variable:", input$welch_group), style = "Normal") %>%
          body_add_par("Method: Welch ANOVA (robust to unequal variances)", style = "Normal") %>%
          body_add_par("", style = "Normal")
        
        doc <- doc %>%
          body_add_par("WELCH ANOVA RESULTS", style = "heading 2")
        
        welch_df <- data.frame(
          Statistic = c("F value", "Numerator DF", "Denominator DF", "P-value"),
          Value = c(
            round(results$welch_test$statistic, 4),
            round(results$welch_test$parameter[1], 2),
            round(results$welch_test$parameter[2], 2),
            round(results$welch_test$p.value, 4)
          )
        )
        
        ft_welch <- flextable(welch_df) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_welch)
        
        if (!is.null(results$posthoc)) {
          doc <- doc %>%
            body_add_par("GAMES-HOWELL POST-HOC TEST", style = "heading 2")
          
          posthoc_text <- capture.output(print(results$posthoc))
          for (line in posthoc_text) {
            doc <- doc %>% body_add_par(line, style = "Normal")
          }
        }
        
        doc <- doc %>%
          body_add_par("INTERPRETATION", style = "heading 2")
        
        p_value <- results$welch_test$p.value
        f_value <- results$welch_test$statistic
        
        interpretation <- if (p_value < 0.05) {
          paste("The Welch ANOVA revealed a statistically significant difference between groups ",
                "(F =", round(f_value, 3), ", p =", round(p_value, 4), "). ",
                "This robust test confirms significant differences while accounting for potential unequal variances.")
        } else {
          paste("The Welch ANOVA did not reveal a statistically significant difference between groups ",
                "(F =", round(f_value, 3), ", p =", round(p_value, 4), ").")
        }
        
        doc <- doc %>% body_add_par(interpretation, style = "Normal")
      }
      
      print(doc, target = file)
    }
  )
  
  # One-Sample T-Test Download
  output$download_ttest_onesample <- downloadHandler(
    filename = function() {
      paste("onesample_ttest_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(ttest_onesample_results())
      results <- ttest_onesample_results()
      
      doc <- read_docx()
      
      if (!results$success) {
        doc <- doc %>%
          body_add_par("ONE-SAMPLE T-TEST - ERROR", style = "heading 1") %>%
          body_add_par(results$error, style = "Normal")
      } else {
        doc <- doc %>%
          body_add_par("ONE-SAMPLE T-TEST ANALYSIS REPORT", style = "heading 1") %>%
          body_add_par(paste("Variable:", input$ttest_onesample_var), style = "Normal") %>%
          body_add_par(paste("Test Value (μ₀):", input$ttest_onesample_mu), style = "Normal") %>%
          body_add_par(paste("Alternative Hypothesis:", input$ttest_onesample_alternative), style = "Normal") %>%
          body_add_par(paste("Confidence Level:", input$ttest_onesample_conf * 100, "%"), style = "Normal") %>%
          body_add_par("", style = "Normal")
        
        doc <- doc %>%
          body_add_par("T-TEST RESULTS", style = "heading 2")
        
        ttest_df <- data.frame(
          Statistic = c("t value", "Degrees of Freedom", "P-value", 
                        "Sample Mean", "Test Value", "Mean Difference",
                        "95% CI Lower", "95% CI Upper"),
          Value = c(
            round(results$ttest$statistic, 4),
            round(results$ttest$parameter, 2),
            round(results$ttest$p.value, 4),
            round(mean(results$data), 4),
            input$ttest_onesample_mu,
            round(mean(results$data) - input$ttest_onesample_mu, 4),
            round(results$ttest$conf.int[1], 4),
            round(results$ttest$conf.int[2], 4)
          )
        )
        
        ft_ttest <- flextable(ttest_df) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_ttest)
        
        doc <- doc %>%
          body_add_par("DESCRIPTIVE STATISTICS", style = "heading 2")
        
        desc_df <- data.frame(
          Statistic = c("N", "Mean", "Standard Deviation", "Standard Error", 
                        "Minimum", "Maximum", "Range"),
          Value = c(
            results$descriptives$N,
            round(results$descriptives$Mean, 4),
            round(results$descriptives$SD, 4),
            round(results$descriptives$SE, 4),
            round(min(results$data), 4),
            round(max(results$data), 4),
            round(max(results$data) - min(results$data), 4)
          )
        )
        
        ft_desc <- flextable(desc_df) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_desc)
        
        doc <- doc %>%
          body_add_par("INTERPRETATION", style = "heading 2")
        
        p_value <- results$ttest$p.value
        mean_diff <- mean(results$data) - input$ttest_onesample_mu
        
        interpretation <- if (p_value < 0.05) {
          paste("The one-sample t-test revealed a statistically significant difference ",
                "between the sample mean (", round(mean(results$data), 3), ") ",
                "and the test value (", input$ttest_onesample_mu, "), ",
                "t(", round(results$ttest$parameter, 1), ") = ", round(results$ttest$statistic, 3), 
                ", p = ", round(p_value, 4), ". ",
                "The mean difference is ", round(mean_diff, 3), ".",
                sep = "")
        } else {
          paste("The one-sample t-test did not reveal a statistically significant difference ",
                "between the sample mean (", round(mean(results$data), 3), ") ",
                "and the test value (", input$ttest_onesample_mu, "), ",
                "t(", round(results$ttest$parameter, 1), ") = ", round(results$ttest$statistic, 3), 
                ", p = ", round(p_value, 4), ". ",
                "The mean difference is ", round(mean_diff, 3), ".",
                sep = "")
        }
        
        doc <- doc %>% body_add_par(interpretation, style = "Normal")
      }
      
      print(doc, target = file)
    }
  )
  
  # Independent T-Test Download
  output$download_ttest_independent <- downloadHandler(
    filename = function() {
      paste("independent_ttest_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(ttest_independent_results())
      results <- ttest_independent_results()
      
      doc <- read_docx()
      
      if (!results$success) {
        doc <- doc %>%
          body_add_par("INDEPENDENT T-TEST - ERROR", style = "heading 1") %>%
          body_add_par(results$error, style = "Normal")
      } else {
        doc <- doc %>%
          body_add_par("INDEPENDENT T-TEST ANALYSIS REPORT", style = "heading 1") %>%
          body_add_par(paste("Outcome Variable:", input$ttest_independent_var), style = "Normal") %>%
          body_add_par(paste("Grouping Variable:", input$ttest_independent_group), style = "Normal") %>%
          body_add_par(paste("Group 1:", input$ttest_group1), style = "Normal") %>%
          body_add_par(paste("Group 2:", input$ttest_group2), style = "Normal") %>%
          body_add_par(paste("Equal Variances Assumed:", ifelse(input$ttest_var_equal, "Yes", "No")), style = "Normal") %>%
          body_add_par(paste("Alternative Hypothesis:", input$ttest_alternative), style = "Normal") %>%
          body_add_par("", style = "Normal")
        
        doc <- doc %>%
          body_add_par("T-TEST RESULTS", style = "heading 2")
        
        ttest_df <- data.frame(
          Statistic = c("t value", "Degrees of Freedom", "P-value", 
                        "Mean Difference", "Standard Error", "95% CI Lower", "95% CI Upper"),
          Value = c(
            round(results$ttest$statistic, 4),
            round(results$ttest$parameter, 2),
            round(results$ttest$p.value, 4),
            round(results$ttest$estimate[1] - results$ttest$estimate[2], 4),
            round(results$ttest$stderr, 4),
            round(results$ttest$conf.int[1], 4),
            round(results$ttest$conf.int[2], 4)
          )
        )
        
        ft_ttest <- flextable(ttest_df) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_ttest)
        
        doc <- doc %>%
          body_add_par("GROUP DESCRIPTIVES", style = "heading 2")
        
        ft_desc <- flextable(results$descriptives) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_desc)
        
        doc <- doc %>%
          body_add_par("VARIANCE EQUALITY TEST", style = "heading 2")
        
        var_df <- data.frame(
          Statistic = c("F value", "Numerator DF", "Denominator DF", "P-value", "Interpretation"),
          Value = c(
            round(results$var_test$statistic, 4),
            round(results$var_test$parameter[1], 2),
            round(results$var_test$parameter[2], 2),
            round(results$var_test$p.value, 4),
            ifelse(results$var_test$p.value < 0.05, 
                   "Variances significantly different - Welch test recommended", 
                   "Variances not significantly different - equal variances assumption reasonable")
          )
        )
        
        ft_var <- flextable(var_df) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_var)
        
        doc <- doc %>%
          body_add_par("INTERPRETATION", style = "heading 2")
        
        p_value <- results$ttest$p.value
        mean_diff <- results$ttest$estimate[1] - results$ttest$estimate[2]
        group1_mean <- results$descriptives$Mean[1]
        group2_mean <- results$descriptives$Mean[2]
        
        interpretation <- if (p_value < 0.05) {
          paste("The independent t-test revealed a statistically significant difference ",
                "between ", input$ttest_group1, " (M = ", round(group1_mean, 3), ") ",
                "and ", input$ttest_group2, " (M = ", round(group2_mean, 3), "), ",
                "t(", round(results$ttest$parameter, 1), ") = ", round(abs(results$ttest$statistic), 3), 
                ", p = ", round(p_value, 4), ". ",
                "The mean difference is ", round(mean_diff, 3), " ",
                "with a ", input$ttest_conf * 100, "% CI of [", 
                round(results$ttest$conf.int[1], 3), ", ", 
                round(results$ttest$conf.int[2], 3), "].",
                sep = "")
        } else {
          paste("The independent t-test did not reveal a statistically significant difference ",
                "between ", input$ttest_group1, " (M = ", round(group1_mean, 3), ") ",
                "and ", input$ttest_group2, " (M = ", round(group2_mean, 3), "), ",
                "t(", round(results$ttest$parameter, 1), ") = ", round(abs(results$ttest$statistic), 3), 
                ", p = ", round(p_value, 4), ". ",
                "The mean difference is ", round(mean_diff, 3), " ",
                "with a ", input$ttest_conf * 100, "% CI of [", 
                round(results$ttest$conf.int[1], 3), ", ", 
                round(results$ttest$conf.int[2], 3), "].",
                sep = "")
        }
        
        doc <- doc %>% body_add_par(interpretation, style = "Normal")
      }
      
      print(doc, target = file)
    }
  )
  
  # Paired T-Test Download
  output$download_ttest_paired <- downloadHandler(
    filename = function() {
      paste("paired_ttest_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(ttest_paired_results())
      results <- ttest_paired_results()
      
      doc <- read_docx()
      
      if (!results$success) {
        doc <- doc %>%
          body_add_par("PAIRED T-TEST - ERROR", style = "heading 1") %>%
          body_add_par(results$error, style = "Normal")
      } else {
        doc <- doc %>%
          body_add_par("PAIRED T-TEST ANALYSIS REPORT", style = "heading 1") %>%
          body_add_par(paste("Variable 1 (Before/Time 1):", input$ttest_paired_var1), style = "Normal") %>%
          body_add_par(paste("Variable 2 (After/Time 2):", input$ttest_paired_var2), style = "Normal") %>%
          body_add_par(paste("Number of Pairs:", length(results$var1_data)), style = "Normal") %>%
          body_add_par(paste("Alternative Hypothesis:", input$ttest_paired_alternative), style = "Normal") %>%
          body_add_par("", style = "Normal")
        
        doc <- doc %>%
          body_add_par("PAIRED T-TEST RESULTS", style = "heading 2")
        
        ttest_df <- data.frame(
          Statistic = c("t value", "Degrees of Freedom", "P-value", 
                        "Mean Difference", "Standard Error", "95% CI Lower", "95% CI Upper"),
          Value = c(
            round(results$ttest$statistic, 4),
            round(results$ttest$parameter, 2),
            round(results$ttest$p.value, 4),
            round(results$ttest$estimate, 4),
            round(results$ttest$stderr, 4),
            round(results$ttest$conf.int[1], 4),
            round(results$ttest$conf.int[2], 4)
          )
        )
        
        ft_ttest <- flextable(ttest_df) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_ttest)
        
        doc <- doc %>%
          body_add_par("DESCRIPTIVE STATISTICS", style = "heading 2")
        
        ft_desc <- flextable(results$descriptives) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_desc)
        
        doc <- doc %>%
          body_add_par("DIFFERENCE ANALYSIS", style = "heading 2")
        
        correlation <- cor(results$var1_data, results$var2_data)
        diff_df <- data.frame(
          Statistic = c("Correlation between pairs", "Mean of differences", 
                        "SD of differences", "Minimum difference", "Maximum difference"),
          Value = c(
            round(correlation, 4),
            round(mean(results$differences), 4),
            round(sd(results$differences), 4),
            round(min(results$differences), 4),
            round(max(results$differences), 4)
          )
        )
        
        ft_diff <- flextable(diff_df) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_diff)
        
        doc <- doc %>%
          body_add_par("INTERPRETATION", style = "heading 2")
        
        p_value <- results$ttest$p.value
        mean_diff <- results$ttest$estimate
        var1_mean <- results$descriptives$Mean[1]
        var2_mean <- results$descriptives$Mean[2]
        
        interpretation <- if (p_value < 0.05) {
          paste("The paired t-test revealed a statistically significant difference ",
                "between ", input$ttest_paired_var1, " (M = ", round(var1_mean, 3), ") ",
                "and ", input$ttest_paired_var2, " (M = ", round(var2_mean, 3), "), ",
                "t(", round(results$ttest$parameter, 1), ") = ", round(abs(results$ttest$statistic), 3), 
                ", p = ", round(p_value, 4), ". ",
                "The mean difference is ", round(mean_diff, 3), " ",
                "with a correlation of ", round(correlation, 3), " between the paired measurements.",
                sep = "")
        } else {
          paste("The paired t-test did not reveal a statistically significant difference ",
                "between ", input$ttest_paired_var1, " (M = ", round(var1_mean, 3), ") ",
                "and ", input$ttest_paired_var2, " (M = ", round(var2_mean, 3), "), ",
                "t(", round(results$ttest$parameter, 1), ") = ", round(abs(results$ttest$statistic), 3), 
                ", p = ", round(p_value, 4), ". ",
                "The mean difference is ", round(mean_diff, 3), " ",
                "with a correlation of ", round(correlation, 3), " between the paired measurements.",
                sep = "")
        }
        
        doc <- doc %>% body_add_par(interpretation, style = "Normal")
      }
      
      print(doc, target = file)
    }
  )
  
  # Non-parametric Tests Download
  output$download_nonparametric <- downloadHandler(
    filename = function() {
      paste("nonparametric_tests_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(mannwhitney_results(), wilcoxon_results())
      
      doc <- read_docx() %>%
        body_add_par("NON-PARAMETRIC TESTS REPORT", style = "heading 1") %>%
        body_add_par("", style = "Normal")
      
      mann_results <- mannwhitney_results()
      if (mann_results$success) {
        doc <- doc %>%
          body_add_par("MANN-WHITNEY U TEST", style = "heading 2") %>%
          body_add_par(paste("Variable:", input$mannwhitney_var), style = "Normal") %>%
          body_add_par(paste("Grouping Variable:", input$mannwhitney_group), style = "Normal") %>%
          body_add_par("", style = "Normal")
        
        mann_df <- data.frame(
          Statistic = c("W statistic", "P-value", "Alternative Hypothesis"),
          Value = c(
            round(mann_results$test$statistic, 4),
            round(mann_results$test$p.value, 4),
            mann_results$test$alternative
          )
        )
        
        ft_mann <- flextable(mann_df) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_mann)
      }
      
      wilcox_results <- wilcoxon_results()
      if (wilcox_results$success) {
        doc <- doc %>%
          body_add_par("WILCOXON SIGNED-RANK TEST", style = "heading 2") %>%
          body_add_par(paste("Variable 1:", input$wilcoxon_var1), style = "Normal") %>%
          body_add_par(paste("Variable 2:", input$wilcoxon_var2), style = "Normal") %>%
          body_add_par("", style = "Normal")
        
        wilcox_df <- data.frame(
          Statistic = c("V statistic", "P-value", "Alternative Hypothesis"),
          Value = c(
            round(wilcox_results$test$statistic, 4),
            round(wilcox_results$test$p.value, 4),
            wilcox_results$test$alternative
          )
        )
        
        ft_wilcox <- flextable(wilcox_df) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_wilcox)
      }
      
      print(doc, target = file)
    }
  )
  
  # ANOVA Assumptions Download - CORRECTED VERSION
  output$download_assumptions <- downloadHandler(
    filename = function() {
      paste("anova_assumptions_report_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(assumptions_results())
      
      results <- assumptions_results()
      
      # Create a new Word document
      doc <- read_docx()
      
      # Title
      doc <- doc %>%
        body_add_par("ANOVA ASSUMPTIONS CHECK REPORT", style = "heading 1") %>%
        body_add_par("", style = "Normal") %>%
        body_add_par(paste("Analysis Date:", Sys.Date()), style = "Normal") %>%
        body_add_par(paste("Outcome Variable:", input$assumptions_outcome), style = "Normal") %>%
        body_add_par(paste("Grouping Variable:", input$assumptions_group), style = "Normal") %>%
        body_add_par("", style = "Normal")
      
      # Error section if analysis failed
      if (!is.null(results$error)) {
        doc <- doc %>%
          body_add_par("ANALYSIS ERROR", style = "heading 2") %>%
          body_add_par(results$error, style = "Normal")
      } else {
        # Levene's Test section
        if (input$levene_test && !is.null(results$levene)) {
          doc <- doc %>%
            body_add_par("HOMOGENEITY OF VARIANCE", style = "heading 2") %>%
            body_add_par("Levene's Test Results:", style = "Normal")
          
          # Format Levene's test results
          levene_df <- as.data.frame(results$levene)
          levene_df <- cbind(Source = rownames(levene_df), levene_df)
          rownames(levene_df) <- NULL
          
          # Clean column names
          colnames(levene_df) <- c("Source", "DF1", "DF2", "F Value", "P Value")
          
          ft_levene <- flextable(levene_df) %>%
            theme_zebra() %>%
            autofit()
          
          doc <- doc %>% body_add_flextable(ft_levene)
          
          # Interpretation
          p_value <- results$levene$`Pr(>F)`[1]
          interpretation <- if (p_value > 0.05) {
            "✓ Assumption met: Variances are homogeneous across groups (p > 0.05)"
          } else {
            paste("✗ Assumption violated: Variances are not homogeneous across groups (p =", 
                  round(p_value, 4), ")")
          }
          
          doc <- doc %>%
            body_add_par("", style = "Normal") %>%
            body_add_par(interpretation, style = "Normal") %>%
            body_add_par("", style = "Normal")
        }
        
        # Normality Tests section
        if (input$normality_test && !is.null(results$shapiro)) {
          doc <- doc %>%
            body_add_par("NORMALITY TESTS", style = "heading 2")
          
          # Create normality test results table
          normality_df <- data.frame(
            Test = character(),
            Statistic = character(),
            P_Value = character(),
            Interpretation = character(),
            stringsAsFactors = FALSE
          )
          
          # Shapiro-Wilk test
          normality_df <- rbind(normality_df, data.frame(
            Test = "Shapiro-Wilk",
            Statistic = round(results$shapiro$statistic, 4),
            P_Value = round(results$shapiro$p.value, 4),
            Interpretation = ifelse(results$shapiro$p.value > 0.05, 
                                    "✓ Assumption met", 
                                    "✗ Assumption violated")
          ))
          
          # Kolmogorov-Smirnov test if available
          if (!is.null(results$ks)) {
            normality_df <- rbind(normality_df, data.frame(
              Test = "Kolmogorov-Smirnov",
              Statistic = round(results$ks$statistic, 4),
              P_Value = round(results$ks$p.value, 4),
              Interpretation = ifelse(results$ks$p.value > 0.05, 
                                      "✓ Assumption met", 
                                      "✗ Assumption violated")
            ))
          }
          
          ft_normality <- flextable(normality_df) %>%
            theme_zebra() %>%
            autofit()
          
          doc <- doc %>% body_add_flextable(ft_normality)
          
          doc <- doc %>%
            body_add_par("", style = "Normal") %>%
            body_add_par("Interpretation Guide:", style = "Normal") %>%
            body_add_par("• P > 0.05: Residuals are normally distributed", style = "Normal") %>%
            body_add_par("• P ≤ 0.05: Residuals are not normally distributed", style = "Normal") %>%
            body_add_par("", style = "Normal")
        }
        
        # Outlier Detection section
        if (input$outlier_test && !is.null(results$outliers)) {
          doc <- doc %>%
            body_add_par("OUTLIER DETECTION", style = "heading 2")
          
          if (length(results$outliers) > 0) {
            outlier_text <- paste("Outliers detected at observation numbers:", 
                                  paste(results$outliers, collapse = ", "))
            outlier_count <- length(results$outliers)
            total_obs <- nrow(processed_data())
            percentage <- round((outlier_count / total_obs) * 100, 2)
            
            doc <- doc %>%
              body_add_par(outlier_text, style = "Normal") %>%
              body_add_par(paste("Total outliers:", outlier_count, "out of", total_obs, 
                                 "observations (", percentage, "%)"), style = "Normal")
            
            # Create outlier summary table
            outlier_summary <- data.frame(
              Metric = c("Number of Outliers", "Total Observations", "Percentage", 
                         "Threshold (SD > |2.5|)"),
              Value = c(outlier_count, total_obs, 
                        paste0(percentage, "%"), "2.5")
            )
            
            ft_outliers <- flextable(outlier_summary) %>%
              theme_zebra() %>%
              autofit()
            
            doc <- doc %>% body_add_flextable(ft_outliers)
            
            doc <- doc %>%
              body_add_par("", style = "Normal") %>%
              body_add_par("Recommendations:", style = "Normal") %>%
              body_add_par("• Review outliers for data entry errors", style = "Normal") %>%
              body_add_par("• Consider robust ANOVA methods if outliers are valid", style = "Normal") %>%
              body_add_par("• Transform data if outliers affect normality", style = "Normal")
          } else {
            doc <- doc %>%
              body_add_par("✓ No significant outliers detected (standardized residuals < |2.5|)", 
                           style = "Normal")
          }
          doc <- doc %>% body_add_par("", style = "Normal")
        }
        
        # Overall Assessment section
        doc <- doc %>%
          body_add_par("OVERALL ASSESSMENT", style = "heading 2")
        
        # Collect issues
        issues <- character()
        recommendations <- character()
        
        # Check Levene's test
        if (input$levene_test && !is.null(results$levene)) {
          if (results$levene$`Pr(>F)`[1] < 0.05) {
            issues <- c(issues, "Heterogeneous variances (p < 0.05)")
            recommendations <- c(recommendations, "Use Welch ANOVA or data transformation")
          }
        }
        
        # Check normality
        if (input$normality_test && !is.null(results$shapiro)) {
          if (results$shapiro$p.value < 0.05) {
            issues <- c(issues, "Non-normal residuals (p < 0.05)")
            recommendations <- c(recommendations, "Use non-parametric tests or data transformation")
          }
        }
        
        # Check outliers
        if (input$outlier_test && !is.null(results$outliers)) {
          if (length(results$outliers) > 0) {
            outlier_count <- length(results$outliers)
            issues <- c(issues, paste(outlier_count, "outliers detected"))
            recommendations <- c(recommendations, "Review outliers or use robust methods")
          }
        }
        
        # Create assessment
        if (length(issues) == 0) {
          assessment <- "All key ANOVA assumptions appear to be reasonably met. Results from standard ANOVA should be valid."
        } else {
          assessment <- paste("Potential issues detected:", paste(issues, collapse = "; "))
        }
        
        doc <- doc %>%
          body_add_par(assessment, style = "Normal") %>%
          body_add_par("", style = "Normal")
        
        # Add recommendations if any issues
        if (length(recommendations) > 0) {
          doc <- doc %>%
            body_add_par("RECOMMENDATIONS", style = "heading 2")
          
          for (rec in unique(recommendations)) {
            doc <- doc %>% body_add_par(paste("•", rec), style = "Normal")
          }
        }
        
        # Add statistical notes
        doc <- doc %>%
          body_add_par("", style = "Normal") %>%
          body_add_par("STATISTICAL NOTES", style = "heading 2") %>%
          body_add_par("ANOVA Assumptions:", style = "Normal") %>%
          body_add_par("1. Independence: Observations are independent", style = "Normal") %>%
          body_add_par("2. Normality: Residuals are normally distributed", style = "Normal") %>%
          body_add_par("3. Homogeneity of variance: Equal variances across groups", style = "Normal") %>%
          body_add_par("4. No influential outliers", style = "Normal") %>%
          body_add_par("", style = "Normal") %>%
          body_add_par("Robust Alternatives:", style = "Normal") %>%
          body_add_par("• Welch ANOVA: For unequal variances", style = "Normal") %>%
          body_add_par("• Kruskal-Wallis test: For non-normal data", style = "Normal") %>%
          body_add_par("• Transformations: Log, square root, or rank transformations", style = "Normal")
      }
      
      # Add footer
      doc <- doc %>%
        body_add_par("", style = "Normal") %>%
        body_add_par("", style = "Normal") %>%
        body_add_par("Report generated by EpiDem Suite™", style = "Normal") %>%
        body_add_par("Epidemiological Analysis Tool", style = "Normal") %>%
        body_add_par(paste("Generated on:", Sys.time()), style = "Normal")
      
      # Save the document
      print(doc, target = file)
    }
  )
  # Server function - CORRECTED RISK RATIO SECTION
  
  # UI for outcome and exposure level selection
  output$rr_outcome_levels_ui <- renderUI({
    req(input$rr_outcome, processed_data())
    df <- processed_data()
    
    if (!input$rr_outcome %in% names(df)) return(NULL)
    
    var_data <- df[[input$rr_outcome]]
    levels <- if (is.factor(var_data)) levels(var_data) else unique(na.omit(var_data))
    
    if (length(levels) > 10) {
      return(helpText("Outcome variable has too many categories for risk ratio calculation."))
    }
    
    tagList(
      selectInput("rr_outcome_ref", "Outcome Reference Level (Baseline):", choices = levels),
      selectInput("rr_outcome_comp", "Outcome Comparison Level:", choices = levels, selected = levels[2])
    )
  })
  
  output$rr_exposure_levels_ui <- renderUI({
    req(input$rr_exposure, processed_data())
    df <- processed_data()
    
    if (!input$rr_exposure %in% names(df)) return(NULL)
    
    var_data <- df[[input$rr_exposure]]
    levels <- if (is.factor(var_data)) levels(var_data) else unique(na.omit(var_data))
    
    if (length(levels) > 10) {
      return(helpText("Exposure variable has too many categories for risk ratio calculation."))
    }
    
    tagList(
      selectInput("rr_exposure_ref", "Exposure Reference Level (Unexposed):", choices = levels),
      selectInput("rr_exposure_comp", "Exposure Comparison Level (Exposed):", choices = levels, selected = levels[2])
    )
  })
  
  # ULTRA-ROBUST RISK RATIO ANALYSIS - COMPLETELY REWRITTEN
  rr_results <- eventReactive(input$run_rr, {
    req(input$rr_outcome, input$rr_exposure, processed_data())
    req(input$rr_outcome_ref, input$rr_outcome_comp, input$rr_exposure_ref, input$rr_exposure_comp)
    
    tryCatch({
      df <- processed_data()
      
      # Validate that variables exist
      if (!input$rr_outcome %in% names(df) || !input$rr_exposure %in% names(df)) {
        return(list(
          error = "Selected variables not found in dataset",
          success = FALSE
        ))
      }
      
      # Get the variables
      outcome_var <- df[[input$rr_outcome]]
      exposure_var <- df[[input$rr_exposure]]
      
      # Remove missing values
      complete_cases <- !is.na(outcome_var) & !is.na(exposure_var)
      outcome_clean <- outcome_var[complete_cases]
      exposure_clean <- exposure_var[complete_cases]
      
      if (length(outcome_clean) == 0 || length(exposure_clean) == 0) {
        return(list(
          error = "No complete cases after removing missing values",
          success = FALSE
        ))
      }
      
      # Convert to factors with user-specified levels
      outcome_factor <- factor(outcome_clean)
      exposure_factor <- factor(exposure_clean)
      
      # Ensure user-selected levels exist in the data
      if (!input$rr_outcome_ref %in% levels(outcome_factor) || 
          !input$rr_outcome_comp %in% levels(outcome_factor)) {
        return(list(
          error = "Selected outcome levels not found in data",
          success = FALSE
        ))
      }
      
      if (!input$rr_exposure_ref %in% levels(exposure_factor) || 
          !input$rr_exposure_comp %in% levels(exposure_factor)) {
        return(list(
          error = "Selected exposure levels not found in data",
          success = FALSE
        ))
      }
      
      # Relevel factors based on user selection
      outcome_factor <- relevel(outcome_factor, ref = input$rr_outcome_ref)
      exposure_factor <- relevel(exposure_factor, ref = input$rr_exposure_ref)
      
      # Create 2x2 contingency table
      tab <- table(
        Exposure = exposure_factor,
        Outcome = outcome_factor
      )
      
      # Keep only the two selected levels for each variable
      tab_2x2 <- tab[
        c(input$rr_exposure_ref, input$rr_exposure_comp),
        c(input$rr_outcome_ref, input$rr_outcome_comp)
      ]
      
      # Check for zero cells
      if (any(tab_2x2 == 0)) {
        tab_2x2_corrected <- tab_2x2 + 0.5
        correction_applied <- TRUE
      } else {
        tab_2x2_corrected <- tab_2x2
        correction_applied <- FALSE
      }
      
      # Extract values from the 2x2 table
      a <- tab_2x2_corrected[2, 2]  # Exposed cases
      b <- tab_2x2_corrected[2, 1]  # Exposed non-cases
      c <- tab_2x2_corrected[1, 2]  # Unexposed cases
      d <- tab_2x2_corrected[1, 1]  # Unexposed non-cases
      
      # Calculate risks
      risk_exposed <- a / (a + b)
      risk_unexposed <- c / (c + d)
      
      # Calculate risk ratio
      risk_ratio <- risk_exposed / risk_unexposed
      
      # Calculate confidence intervals using log method
      log_rr <- log(risk_ratio)
      se_log_rr <- sqrt((1/a) - (1/(a+b)) + (1/c) - (1/(c+d)))
      
      z <- qnorm(1 - (1 - input$rr_conf) / 2)
      ci_lower <- exp(log_rr - z * se_log_rr)
      ci_upper <- exp(log_rr + z * se_log_rr)
      
      # Calculate risk difference
      risk_diff <- risk_exposed - risk_unexposed
      se_rd <- sqrt((risk_exposed * (1 - risk_exposed) / (a + b)) + 
                      (risk_unexposed * (1 - risk_unexposed) / (c + d)))
      rd_ci_lower <- risk_diff - z * se_rd
      rd_ci_upper <- risk_diff + z * se_rd
      
      # Create comprehensive results table
      results_table <- data.frame(
        Statistic = c(
          "Exposed Group",
          "Unexposed Group", 
          "Risk in Exposed",
          "Risk in Unexposed",
          "Risk Ratio (RR)",
          paste("RR Lower CI (", input$rr_conf * 100, "%)", sep = ""),
          paste("RR Upper CI (", input$rr_conf * 100, "%)", sep = ""),
          "Risk Difference",
          paste("RD Lower CI (", input$rr_conf * 100, "%)", sep = ""),
          paste("RD Upper CI (", input$rr_conf * 100, "%)", sep = ""),
          "Exposed Cases",
          "Exposed Non-Cases",
          "Unexposed Cases", 
          "Unexposed Non-Cases",
          "Total Exposed",
          "Total Unexposed",
          "Total Cases",
          "Total Observations",
          "Zero Cell Correction"
        ),
        Value = c(
          input$rr_exposure_comp,
          input$rr_exposure_ref,
          round(risk_exposed, 4),
          round(risk_unexposed, 4),
          round(risk_ratio, 4),
          round(ci_lower, 4),
          round(ci_upper, 4),
          round(risk_diff, 4),
          round(rd_ci_lower, 4),
          round(rd_ci_upper, 4),
          a, b, c, d,
          a + b,
          c + d,
          a + c,
          a + b + c + d,
          ifelse(correction_applied, "Yes (Haldane-Anscombe)", "No")
        ),
        stringsAsFactors = FALSE
      )
      
      # Create contingency table for display
      contingency_df <- as.data.frame.matrix(tab_2x2)
      colnames(contingency_df) <- paste("Outcome:", colnames(contingency_df))
      rownames(contingency_df) <- paste("Exposure:", rownames(contingency_df))
      
      # Add totals
      contingency_df$Total <- rowSums(contingency_df)
      total_row <- c(colSums(contingency_df[, -ncol(contingency_df), drop = FALSE]), 
                     sum(contingency_df$Total))
      contingency_df <- rbind(contingency_df, Total = total_row)
      
      list(
        table = results_table,
        contingency = contingency_df,
        raw_table = tab_2x2,
        risk_ratio = risk_ratio,
        ci_lower = ci_lower,
        ci_upper = ci_upper,
        risk_exposed = risk_exposed,
        risk_unexposed = risk_unexposed,
        correction_applied = correction_applied,
        success = TRUE
      )
      
    }, error = function(e) {
      return(list(
        error = paste("Risk ratio calculation failed:", e$message),
        success = FALSE
      ))
    })
  })
  
  # Output for risk ratio results
  output$rr_results <- renderPrint({
    req(rr_results())
    
    results <- rr_results()
    
    if (!results$success) {
      cat("RISK RATIO ANALYSIS - ERROR\n")
      cat("===========================\n")
      cat(results$error, "\n")
      return()
    }
    
    cat("RISK RATIO ANALYSIS\n")
    cat("===================\n\n")
    
    cat("CONTINGENCY TABLE:\n")
    print(results$raw_table)
    cat("\n")
    
    cat("RISK RATIO RESULTS:\n")
    print(results$table)
  })
  
  # Additional interpretation output
  output$rr_interpretation <- renderPrint({
    req(rr_results())
    
    results <- rr_results()
    
    if (!results$success) return()
    
    cat("\nINTERPRETATION:\n")
    cat("===============\n\n")
    
    rr <- results$risk_ratio
    ci_low <- results$ci_lower
    ci_high <- results$ci_upper
    
    if (rr > 1) {
      cat(sprintf("• The %s group has %.2f times the risk of %s compared to the %s group.\n", 
                  input$rr_exposure_comp, rr, input$rr_outcome_comp, input$rr_exposure_ref))
    } else if (rr < 1) {
      cat(sprintf("• The %s group has %.2f times the risk of %s compared to the %s group.\n", 
                  input$rr_exposure_comp, rr, input$rr_outcome_comp, input$rr_exposure_ref))
      cat(sprintf("  (This represents a %.1f%% reduction in risk)\n", (1 - rr) * 100))
    } else {
      cat(sprintf("• There is no difference in risk between the %s and %s groups.\n", 
                  input$rr_exposure_comp, input$rr_exposure_ref))
    }
    
    cat(sprintf("• We are %d%% confident that the true risk ratio lies between %.2f and %.2f.\n", 
                input$rr_conf * 100, ci_low, ci_high))
    
    if (ci_low > 1 || ci_high < 1) {
      cat("• This result is statistically significant (confidence interval does not include 1).\n")
    } else {
      cat("• This result is not statistically significant (confidence interval includes 1).\n")
    }
    
    if (results$correction_applied) {
      cat("• Note: Haldane-Anscombe correction was applied due to zero cells in the contingency table.\n")
    }
    
    # Public health significance
    risk_diff <- results$risk_exposed - results$risk_unexposed
    if (abs(risk_diff) > 0.1) {
      cat(sprintf("• This represents an absolute risk difference of %.1f%%.\n", risk_diff * 100))
    }
  })
  
  output$rr_table <- renderDT({
    req(rr_results())
    
    results <- rr_results()
    
    if (!results$success) {
      datatable(
        data.frame(Error = results$error),
        options = list(dom = 't'),
        rownames = FALSE,
        caption = "Risk Ratio Analysis - Error"
      )
    } else {
      datatable(
        results$table,
        options = list(
          pageLength = 20,
          dom = 'Blfrtip',
          scrollX = TRUE,
          buttons = c('copy', 'csv', 'excel')
        ),
        rownames = FALSE,
        caption = "Risk Ratio Analysis Results",
        extensions = 'Buttons'
      ) %>%
        formatStyle(
          'Value',
          backgroundColor = styleEqual(
            c("Yes (Haldane-Anscombe)"),
            c('#FFF3CD')
          )
        )
    }
  })
  
  # Download handler for risk ratios
  output$download_rr <- downloadHandler(
    filename = function() {
      paste("risk_ratio_analysis_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(rr_results())
      
      results <- rr_results()
      
      if (!results$success) {
        ft <- flextable(data.frame(Error = results$error)) %>%
          set_caption("Risk Ratio Analysis - Error") %>%
          theme_zebra()
        
        doc <- read_docx() %>% 
          body_add_flextable(ft)
      } else {
        doc <- read_docx()
        
        doc <- doc %>%
          body_add_par("RISK RATIO ANALYSIS REPORT", style = "heading 1") %>%
          body_add_par(paste("Outcome Variable:", input$rr_outcome), style = "Normal") %>%
          body_add_par(paste("Exposure Variable:", input$rr_exposure), style = "Normal") %>%
          body_add_par(paste("Confidence Level:", input$rr_conf * 100, "%"), style = "Normal") %>%
          body_add_par("", style = "Normal")
        
        doc <- doc %>%
          body_add_par("CONTINGENCY TABLE", style = "heading 2")
        
        contingency_with_rownames <- cbind(
          Category = rownames(results$contingency), 
          results$contingency
        )
        ft_contingency <- flextable(contingency_with_rownames) %>%
          theme_zebra() %>%
          autofit()
        doc <- doc %>% body_add_flextable(ft_contingency)
        
        doc <- doc %>%
          body_add_par("RISK RATIO RESULTS", style = "heading 2")
        
        ft_results <- flextable(results$table) %>%
          theme_zebra() %>%
          autofit()
        doc <- doc %>% body_add_flextable(ft_results)
        
        doc <- doc %>%
          body_add_par("INTERPRETATION", style = "heading 2")
        
        rr <- results$risk_ratio
        interpretation_text <- if (rr > 1) {
          paste("The", input$rr_exposure_comp, "group has", round(rr, 2), 
                "times the risk of", input$rr_outcome_comp, "compared to the", 
                input$rr_exposure_ref, "group.")
        } else if (rr < 1) {
          paste("The", input$rr_exposure_comp, "group has", round(rr, 2), 
                "times the risk of", input$rr_outcome_comp, "compared to the", 
                input$rr_exposure_ref, "group, representing a", 
                round((1 - rr) * 100, 1), "% reduction in risk.")
        } else {
          paste("There is no difference in risk between the", input$rr_exposure_comp, 
                "and", input$rr_exposure_ref, "groups.")
        }
        
        doc <- doc %>% 
          body_add_par(interpretation_text, style = "Normal") %>%
          body_add_par(paste("Confidence interval:", round(results$ci_lower, 2), "to", 
                             round(results$ci_upper, 2)), style = "Normal")
      }
      
      print(doc, target = file)
    }
  )
  
  # Ultra-robust Odds Ratio Analysis - FIXED for Inf and calculation issues
  or_results <- eventReactive(input$run_or, {
    req(input$or_outcome, input$or_exposure, processed_data())
    
    tryCatch({
      df <- processed_data()
      
      # Double-check variables exist with proper NA handling
      if (is.null(input$or_outcome) || !input$or_outcome %in% names(df)) {
        return(list(
          table = data.frame(),
          contingency = data.frame(),
          error = "Outcome variable not found in dataset",
          success = FALSE
        ))
      }
      
      if (is.null(input$or_exposure) || !input$or_exposure %in% names(df)) {
        return(list(
          table = data.frame(),
          contingency = data.frame(),
          error = "Exposure variable not found in dataset",
          success = FALSE
        ))
      }
      
      # Safely get the variables
      outcome_data <- df[[input$or_outcome]]
      exposure_data <- df[[input$or_exposure]]
      
      # Handle empty or all-NA data
      if (length(outcome_data) == 0 || (all(is.na(outcome_data)) && length(outcome_data) > 0)) {
        return(list(
          table = data.frame(),
          contingency = data.frame(),
          error = "Outcome variable contains no data",
          success = FALSE
        ))
      }
      
      if (length(exposure_data) == 0 || (all(is.na(exposure_data)) && length(exposure_data) > 0)) {
        return(list(
          table = data.frame(),
          contingency = data.frame(),
          error = "Exposure variable contains no data",
          success = FALSE
        ))
      }
      
      # Remove rows with missing values in either variable
      complete_cases <- !is.na(outcome_data) & !is.na(exposure_data)
      
      if (sum(complete_cases, na.rm = TRUE) == 0) {
        return(list(
          table = data.frame(),
          contingency = data.frame(),
          error = "No complete cases after removing missing values",
          success = FALSE
        ))
      }
      
      df_clean <- df[complete_cases, ]
      
      # Convert to factors and ensure proper reference levels
      outcome_factor <- as.factor(df_clean[[input$or_outcome]])
      exposure_factor <- as.factor(df_clean[[input$or_exposure]])
      
      # Get levels for user information
      outcome_levels <- levels(outcome_factor)
      exposure_levels <- levels(exposure_factor)
      
      # Ensure we have at least 2 levels
      if (length(outcome_levels) < 2) {
        return(list(
          table = data.frame(),
          contingency = data.frame(),
          error = "Outcome variable must have at least 2 categories",
          success = FALSE
        ))
      }
      
      if (length(exposure_levels) < 2) {
        return(list(
          table = data.frame(),
          contingency = data.frame(),
          error = "Exposure variable must have at least 2 categories",
          success = FALSE
        ))
      }
      
      # Create contingency table
      contingency_table <- table(exposure_factor, outcome_factor)
      
      # Ensure we're working with a 2x2 table for odds ratio
      if (nrow(contingency_table) > 2) {
        row_sums <- rowSums(contingency_table)
        top_two <- names(sort(row_sums, decreasing = TRUE))[1:2]
        contingency_table <- contingency_table[top_two, ]
        exposure_levels <- top_two
      }
      
      if (ncol(contingency_table) > 2) {
        col_sums <- colSums(contingency_table)
        top_two <- names(sort(col_sums, decreasing = TRUE))[1:2]
        contingency_table <- contingency_table[, top_two]
        outcome_levels <- top_two
      }
      
      # Extract the 2x2 table values
      a <- as.numeric(contingency_table[1, 1])  # Exposed cases
      b <- as.numeric(contingency_table[1, 2])  # Exposed non-cases
      c <- as.numeric(contingency_table[2, 1])  # Unexposed cases
      d <- as.numeric(contingency_table[2, 2])  # Unexposed non-cases
      
      # Check for problematic tables that cause Inf OR
      has_zero_cells <- any(c(a, b, c, d) == 0)
      has_undefined_or <- (b == 0 && c == 0) || (a == 0 && d == 0)
      
      # Calculate odds ratio with robust zero-cell handling
      calculate_or_safe <- function(a, b, c, d, conf_level) {
        
        # Apply Haldane-Anscombe correction for zero cells
        if (has_zero_cells) {
          a_corr <- a + 0.5
          b_corr <- b + 0.5
          c_corr <- c + 0.5
          d_corr <- d + 0.5
          correction_applied <- TRUE
        } else {
          a_corr <- a
          b_corr <- b
          c_corr <- c
          d_corr <- d
          correction_applied = FALSE
        }
        
        # Calculate odds ratio
        if (b_corr == 0 && c_corr == 0) {
          return(list(or = Inf, ci_lower = Inf, ci_upper = Inf, method = "undefined"))
        } else if (b_corr == 0) {
          or_value <- Inf
        } else if (c_corr == 0) {
          or_value <- Inf
        } else {
          or_value <- (a_corr * d_corr) / (b_corr * c_corr)
        }
        
        # Calculate confidence intervals if possible
        if (is.finite(or_value) && or_value > 0 && !is.infinite(or_value)) {
          log_or <- log(or_value)
          se_log_or <- sqrt(1/a_corr + 1/b_corr + 1/c_corr + 1/d_corr)
          z <- qnorm(1 - (1 - conf_level) / 2)
          
          ci_lower <- exp(log_or - z * se_log_or)
          ci_upper <- exp(log_or + z * se_log_or)
        } else {
          ci_lower <- NA
          ci_upper <- NA
        }
        
        return(list(
          or = or_value,
          ci_lower = ci_lower,
          ci_upper = ci_upper,
          method = ifelse(correction_applied, "corrected for zero cells", "standard"),
          correction_applied = correction_applied
        ))
      }
      
      # Calculate odds ratio
      or_calc <- calculate_or_safe(a, b, c, d, input$or_conf)
      
      # Create comprehensive results table
      results_table <- data.frame(
        Statistic = c(
          "Outcome Variable",
          "Outcome Levels",
          "Exposure Variable", 
          "Exposure Levels",
          "Odds Ratio",
          paste("Lower CI (", input$or_conf * 100, "%)", sep = ""),
          paste("Upper CI (", input$or_conf * 100, "%)", sep = ""),
          "Calculation Method",
          "Zero Cells Detected",
          "Correction Applied",
          "Exposed Cases",
          "Exposed Non-Cases", 
          "Unexposed Cases",
          "Unexposed Non-Cases",
          "Total Cases",
          "Total Non-Cases",
          "Total Exposed",
          "Total Unexposed",
          "Total Observations"
        ),
        Value = c(
          input$or_outcome,
          paste(outcome_levels, collapse = ", "),
          input$or_exposure,
          paste(exposure_levels, collapse = ", "),
          ifelse(is.infinite(or_calc$or), "Infinity", 
                 ifelse(is.na(or_calc$or), "Cannot calculate", round(or_calc$or, 4))),
          ifelse(is.na(or_calc$ci_lower), "Cannot calculate", round(or_calc$ci_lower, 4)),
          ifelse(is.na(or_calc$ci_upper), "Cannot calculate", round(or_calc$ci_upper, 4)),
          or_calc$method,
          ifelse(has_zero_cells, "Yes", "No"),
          ifelse(or_calc$correction_applied, "Yes", "No"),
          a, b, c, d,
          a + c,  # Total cases
          b + d,  # Total non-cases
          a + b,  # Total exposed
          c + d,  # Total unexposed
          a + b + c + d  # Total observations
        ),
        stringsAsFactors = FALSE
      )
      
      # Create formatted contingency table for display
      contingency_df <- as.data.frame.matrix(contingency_table)
      colnames(contingency_df) <- paste("Outcome:", colnames(contingency_df))
      rownames(contingency_df) <- paste("Exposure:", rownames(contingency_df))
      
      # Add row and column totals
      contingency_df$Total <- rowSums(contingency_df)
      total_row <- c(colSums(contingency_df[, -ncol(contingency_df), drop = FALSE]), 
                     sum(contingency_df$Total))
      contingency_df <- rbind(contingency_df, Total = total_row)
      
      # Calculate additional statistics
      additional_stats <- data.frame(
        Statistic = character(),
        Value = character(),
        stringsAsFactors = FALSE
      )
      
      # Only calculate chi-square if we have reasonable data
      if (sum(contingency_table) >= 20 && all(contingency_table >= 5)) {
        chi_test <- tryCatch({
          result <- chisq.test(contingency_table)
          data.frame(
            Statistic = c("Chi-square Statistic", "Chi-square P-value"),
            Value = c(round(result$statistic, 4), round(result$p.value, 4))
          )
        }, error = function(e) {
          data.frame(
            Statistic = c("Chi-square Statistic", "Chi-square P-value"),
            Value = c("Cannot calculate", "Cannot calculate")
          )
        })
        additional_stats <- rbind(additional_stats, chi_test)
      }
      
      # Fisher's exact test for small samples or when chi-square assumptions violated
      fisher_test <- tryCatch({
        result <- fisher.test(contingency_table, conf.level = input$or_conf)
        data.frame(
          Statistic = c("Fisher's Exact OR", "Fisher's Exact P-value"),
          Value = c(round(result$estimate, 4), round(result$p.value, 4))
        )
      }, error = function(e) {
        data.frame(
          Statistic = c("Fisher's Exact OR", "Fisher's Exact P-value"),
          Value = c("Cannot calculate", "Cannot calculate")
        )
      })
      additional_stats <- rbind(additional_stats, fisher_test)
      
      # Risk ratio and risk difference
      if (a + b > 0 && c + d > 0) {
        risk_exposed <- a / (a + b)
        risk_unexposed <- c / (c + d)
        risk_ratio <- risk_exposed / risk_unexposed
        risk_diff <- risk_exposed - risk_unexposed
        
        risk_stats <- data.frame(
          Statistic = c("Risk in Exposed", "Risk in Unexposed", "Risk Ratio", "Risk Difference"),
          Value = c(
            round(risk_exposed, 4),
            round(risk_unexposed, 4),
            round(risk_ratio, 4),
            round(risk_diff, 4)
          )
        )
        additional_stats <- rbind(additional_stats, risk_stats)
      }
      
      list(
        table = results_table,
        contingency = contingency_df,
        additional_stats = additional_stats,
        raw_table = contingency_table,
        or_calc = or_calc,
        success = TRUE
      )
      
    }, error = function(e) {
      return(list(
        table = data.frame(),
        contingency = data.frame(),
        error = paste("Error:", e$message),
        success = FALSE
      ))
    })
  })
  
  output$or_results <- renderPrint({
    req(or_results())
    
    results <- or_results()
    
    if (!results$success) {
      cat("ODDS RATIO ANALYSIS - ERROR\n")
      cat("===========================\n")
      cat(results$error, "\n")
      return()
    }
    
    cat("ODDS RATIO ANALYSIS\n")
    cat("===================\n\n")
    
    cat("CONTINGENCY TABLE:\n")
    print(results$raw_table)
    cat("\n")
    
    cat("ODDS RATIO RESULTS:\n")
    print(results$table)
    cat("\n")
    
    if (nrow(results$additional_stats) > 0) {
      cat("ADDITIONAL STATISTICS:\n")
      print(results$additional_stats)
      cat("\n")
    }
    
    # Enhanced interpretation
    or_value <- results$or_calc$or
    ci_lower <- results$or_calc$ci_lower
    ci_upper <- results$or_calc$ci_upper
    
    cat("INTERPRETATION:\n")
    
    if (is.infinite(or_value)) {
      cat("• The odds ratio is infinite due to zero cells in the contingency table.\n")
      cat("• This occurs when there are no cases in one of the exposure groups.\n")
      cat("• Consider using Fisher's exact test or reviewing your data categories.\n")
    } else if (is.na(or_value)) {
      cat("• Cannot calculate odds ratio due to data issues.\n")
    } else {
      if (or_value > 1) {
        cat(sprintf("• The odds of the outcome in the exposed group are %.2f times higher than in the unexposed group.\n", or_value))
      } else if (or_value < 1) {
        cat(sprintf("• The odds of the outcome in the exposed group are %.2f times lower than in the unexposed group.\n", 1/or_value))
      } else {
        cat("• There is no association between exposure and outcome (OR = 1).\n")
      }
      
      if (!is.na(ci_lower) && !is.na(ci_upper)) {
        if (ci_lower > 1 || ci_upper < 1) {
          cat("• The association is statistically significant (confidence interval does not include 1).\n")
        } else {
          cat("• The association is not statistically significant (confidence interval includes 1).\n")
        }
        cat(sprintf("• We are %d%% confident that the true odds ratio lies between %.2f and %.2f.\n", 
                    input$or_conf * 100, ci_lower, ci_upper))
      }
    }
    
    # Data quality notes
    if (results$or_calc$correction_applied) {
      cat("• Note: Haldane-Anscombe correction (adding 0.5 to all cells) was applied due to zero cells.\n")
    }
    
    if (any(results$raw_table == 0)) {
      cat("• Warning: Zero cells detected in the contingency table. Results should be interpreted with caution.\n")
    }
  })
  
  output$or_table <- renderDT({
    req(or_results())
    
    results <- or_results()
    
    if (!results$success) {
      datatable(data.frame(Error = results$error), 
                options = list(dom = 't'),
                rownames = FALSE,
                caption = "Odds Ratio Analysis - Error")
    } else {
      datatable(results$table,
                options = list(
                  pageLength = 15,
                  dom = 'Blfrtip',
                  scrollX = TRUE,
                  buttons = c('copy', 'csv', 'excel')
                ),
                rownames = FALSE,
                caption = "Odds Ratio Analysis Results",
                extensions = 'Buttons'
      ) %>%
        formatStyle(
          'Value',
          backgroundColor = styleEqual(
            c("Infinity", "Cannot calculate", "Yes", "corrected for zero cells"),
            c('#FFE5E5', '#FFE5E5', '#FFF3CD', '#FFF3CD')
          )
        )
    }
  })
  
  output$download_or <- downloadHandler(
    filename = function() {
      paste("odds_ratio_analysis_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(or_results())
      
      results <- or_results()
      
      if (!results$success) {
        ft <- flextable(data.frame(Error = results$error)) %>%
          set_caption("Odds Ratio Analysis - Error") %>%
          theme_zebra()
        
        doc <- read_docx() %>% 
          body_add_flextable(ft)
      } else {
        doc <- read_docx()
        
        doc <- doc %>%
          body_add_par("ODDS RATIO ANALYSIS REPORT", style = "heading 1") %>%
          body_add_par(paste("Outcome:", input$or_outcome), style = "Normal") %>%
          body_add_par(paste("Exposure:", input$or_exposure), style = "Normal") %>%
          body_add_par(paste("Confidence Level:", input$or_conf * 100, "%"), style = "Normal") %>%
          body_add_par("", style = "Normal")
        
        doc <- doc %>%
          body_add_par("CONTINGENCY TABLE", style = "heading 2")
        
        contingency_with_rownames <- cbind(Exposure = rownames(results$contingency), results$contingency)
        ft_contingency <- flextable(contingency_with_rownames) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_contingency)
        
        doc <- doc %>%
          body_add_par("ODDS RATIO RESULTS", style = "heading 2")
        
        ft_main <- flextable(results$table) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_main)
        
        doc <- doc %>%
          body_add_par("ADDITIONAL STATISTICS", style = "heading 2")
        
        ft_additional <- flextable(results$additional_stats) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_additional)
      }
      
      print(doc, target = file)
    }
  )
  
  # Ultra-robust Kaplan-Meier Analysis - FIXED VERSION
  surv_results <- eventReactive(input$run_surv, {
    req(input$surv_time, input$surv_event, processed_data())
    
    tryCatch({
      df <- processed_data()
      
      # Double-check variables exist
      if (!input$surv_time %in% names(df)) {
        return(list(
          fit = NULL,
          error = "Time variable not found in dataset",
          success = FALSE
        ))
      }
      
      if (!input$surv_event %in% names(df)) {
        return(list(
          fit = NULL,
          error = "Event variable not found in dataset",
          success = FALSE
        ))
      }
      
      if (input$surv_group != "none" && !input$surv_group %in% names(df)) {
        return(list(
          fit = NULL,
          error = "Grouping variable not found in dataset",
          success = FALSE
        ))
      }
      
      # Safely get the variables
      time_data <- df[[input$surv_time]]
      event_data <- df[[input$surv_event]]
      
      # Handle empty or all-NA data
      if (length(time_data) == 0 || all(is.na(time_data))) {
        return(list(
          fit = NULL,
          error = "Time variable contains no data",
          success = FALSE
        ))
      }
      
      if (length(event_data) == 0 || all(is.na(event_data))) {
        return(list(
          fit = NULL,
          error = "Event variable contains no data",
          success = FALSE
        ))
      }
      
      # Convert to numeric - ROBUST CONVERSION
      time_numeric <- tryCatch({
        as.numeric(time_data)
      }, warning = function(w) {
        if (is.factor(time_data)) {
          as.numeric(as.character(time_data))
        } else {
          as.numeric(time_data)
        }
      }, error = function(e) {
        rep(NA, length(time_data))
      })
      
      event_numeric <- tryCatch({
        event_num <- as.numeric(event_data)
        if (all(is.na(event_num))) {
          if (is.factor(event_data) || is.character(event_data)) {
            event_num <- as.numeric(as.factor(event_data))
          } else {
            event_num <- as.numeric(event_data)
          }
        }
        event_num
      }, error = function(e) {
        rep(NA, length(event_data))
      })
      
      # Check for successful conversion
      if (all(is.na(time_numeric))) {
        return(list(
          fit = NULL,
          error = "Time variable could not be converted to numeric",
          success = FALSE
        ))
      }
      
      if (all(is.na(event_numeric))) {
        return(list(
          fit = NULL,
          error = "Event variable could not be converted to numeric",
          success = FALSE
        ))
      }
      
      # Remove rows with missing time or event data
      valid_cases <- !is.na(time_numeric) & !is.na(event_numeric) & time_numeric >= 0
      df_clean <- df[valid_cases, ]
      time_clean <- time_numeric[valid_cases]
      event_clean <- event_numeric[valid_cases]
      
      if (nrow(df_clean) == 0) {
        return(list(
          fit = NULL,
          error = "No valid data after removing missing/invalid time and event values",
          success = FALSE
        ))
      }
      
      # Ensure event is binary (0/1) - ROBUST APPROACH
      unique_events <- unique(event_clean)
      if (length(unique_events) == 1) {
        return(list(
          fit = NULL,
          error = "Event variable has only one unique value. Need both events and censoring.",
          success = FALSE
        ))
      }
      
      # Standardize event to 0/1 if needed
      if (!all(unique_events %in% c(0, 1))) {
        event_clean <- ifelse(event_clean == min(unique_events), 0, 1)
      }
      
      # Check if we have both events and censoring
      if (sum(event_clean == 1) == 0) {
        return(list(
          fit = NULL,
          error = "No events found in the data (all observations are censored)",
          success = FALSE
        ))
      }
      
      # Create survival object
      surv_obj <- survival::Surv(time = time_clean, event = event_clean)
      
      # Fit Kaplan-Meier model
      if (input$surv_group != "none") {
        group_data <- df_clean[[input$surv_group]]
        if (all(is.na(group_data))) {
          return(list(
            fit = NULL,
            error = "Grouping variable contains no valid data",
            success = FALSE
          ))
        }
        
        group_var <- as.factor(group_data)
        group_var[is.na(group_var)] <- "Missing"
        
        if (length(levels(group_var)) < 2) {
          return(list(
            fit = NULL,
            error = "Grouping variable must have at least 2 categories",
            success = FALSE
          ))
        }
        
        surv_data <- data.frame(
          time = time_clean,
          event = event_clean,
          group = group_var
        )
        
        surv_fit <- survfit(Surv(time, event) ~ group, data = surv_data)
        
        group_levels <- levels(group_var)
        n_groups <- length(group_levels)
      } else {
        surv_data <- data.frame(
          time = time_clean,
          event = event_clean
        )
        surv_fit <- survfit(Surv(time, event) ~ 1, data = surv_data)
        group_levels <- "Overall"
        n_groups = 1
      }
      
      # Calculate summary statistics
      summary_stats <- summary(surv_fit)
      median_survival <- tryCatch({
        surv_median(surv_fit)
      }, error = function(e) {
        NULL
      })
      
      # Log-rank test if groups exist
      logrank_test <- NULL
      if (input$surv_group != "none" && n_groups > 1) {
        logrank_test <- tryCatch({
          survdiff(Surv(time, event) ~ group, data = surv_data)
        }, error = function(e) {
          NULL
        })
      }
      
      list(
        fit = surv_fit,
        data = surv_data,
        surv_obj = surv_obj,
        group_levels = group_levels,
        n_groups = n_groups,
        summary = summary_stats,
        median_survival = median_survival,
        logrank_test = logrank_test,
        success = TRUE
      )
      
    }, error = function(e) {
      return(list(
        fit = NULL,
        error = paste("Error in Kaplan-Meier analysis:", e$message),
        success = FALSE
      ))
    })
  })
  
  # Reactive for life table times
  life_table_times <- reactive({
    req(input$life_table_times)
    
    times <- strsplit(input$life_table_times, ",")[[1]]
    times <- trimws(times)
    times <- as.numeric(times)
    times <- times[!is.na(times) & times >= 0]
    
    if (length(times) == 0) {
      return(c(0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0))
    }
    
    return(times)
  })
  
  # Life table results - CORRECTED VERSION
  life_table_results <- eventReactive(input$update_life_table, {
    req(surv_results(), input$life_table_times)
    
    results <- surv_results()
    
    if (!results$success) {
      return(list(
        table = data.frame(Error = "No survival analysis results available"),
        summary = "Please run the Kaplan-Meier analysis first.",
        success = FALSE
      ))
    }
    
    tryCatch({
      times <- strsplit(input$life_table_times, ",")[[1]]
      times <- trimws(times)
      times <- as.numeric(times)
      times <- times[!is.na(times) & times >= 0]
      
      if (length(times) == 0) {
        return(list(
          table = data.frame(Error = "No valid time points provided"),
          summary = "Please enter valid numeric time points separated by commas.",
          success = FALSE
        ))
      }
      
      times <- sort(unique(times))
      
      surv_summary <- summary(results$fit, times = times, extend = TRUE)
      
      if (results$n_groups == 1) {
        life_table <- data.frame(
          Time = surv_summary$time,
          Survival_Probability = round(surv_summary$surv, input$life_table_decimals),
          Standard_Error = round(surv_summary$std.err, input$life_table_decimals),
          Lower_CI = round(surv_summary$lower, input$life_table_decimals),
          Upper_CI = round(surv_summary$upper, input$life_table_decimals),
          Number_at_Risk = surv_summary$n.risk,
          Number_of_Events = surv_summary$n.event,
          Number_Censored = surv_summary$n.censor
        )
      } else {
        life_table <- data.frame(
          Group = surv_summary$strata,
          Time = surv_summary$time,
          Survival_Probability = round(surv_summary$surv, input$life_table_decimals),
          Standard_Error = round(surv_summary$std.err, input$life_table_decimals),
          Lower_CI = round(surv_summary$lower, input$life_table_decimals),
          Upper_CI = round(surv_summary$upper, input$life_table_decimals),
          Number_at_Risk = surv_summary$n.risk,
          Number_of_Events = surv_summary$n.event,
          Number_Censored = surv_summary$n.censor
        )
        
        life_table$Group <- gsub(".*=", "", life_table$Group)
      }
      
      total_observations <- nrow(results$data)
      total_events <- sum(results$data$event)
      total_censored <- total_observations - total_events
      
      summary_text <- paste(
        "Life Table Summary:\n",
        "Total Observations: ", total_observations, "\n",
        "Total Events: ", total_events, "\n", 
        "Total Censored: ", total_censored, "\n",
        "Number of Time Points: ", length(times), "\n",
        "Time Range: ", min(times), " to ", max(times), "\n",
        if (results$n_groups > 1) paste("Number of Groups: ", results$n_groups),
        sep = ""
      )
      
      list(
        table = life_table,
        summary = summary_text,
        times = times,
        success = TRUE
      )
      
    }, error = function(e) {
      return(list(
        table = data.frame(Error = paste("Error generating life table:", e$message)),
        summary = "Error in life table generation",
        success = FALSE
      ))
    })
  })
  
  # Life table output - CORRECTED
  output$life_table_output <- renderDT({
    req(life_table_results())
    
    results <- life_table_results()
    
    if (!results$success) {
      datatable(
        results$table,
        options = list(dom = 't', pageLength = 10),
        rownames = FALSE,
        caption = "Life Table - Error"
      )
    } else {
      datatable(
        results$table,
        options = list(
          dom = 'Blfrtip',
          pageLength = 15,
          scrollX = TRUE,
          buttons = c('copy', 'csv', 'excel')
        ),
        rownames = FALSE,
        caption = paste("Life Table - Survival Probabilities at", length(results$times), "Specified Time Points"),
        extensions = 'Buttons'
      )
    }
  })
  
  # Life table summary output - CORRECTED
  output$life_table_summary <- renderPrint({
    req(life_table_results())
    
    results <- life_table_results()
    
    if (!results$success) {
      cat("Life Table Summary - Error\n")
      cat("==========================\n")
      cat(results$table$Error, "\n")
    } else {
      cat(results$summary)
      cat("\n\nSpecified Time Points:\n")
      cat(paste(results$times, collapse = ", "), "\n")
    }
  })
  
  # Download handler for life table (Excel format) - CORRECTED
  output$download_life_table <- downloadHandler(
    filename = function() {
      paste("life_table_", Sys.Date(), ".xlsx", sep = "")
    },
    content = function(file) {
      req(life_table_results())
      
      results <- life_table_results()
      
      if (!results$success) {
        wb <- createWorkbook()
        addWorksheet(wb, "Error")
        writeData(wb, "Error", data.frame(Error = results$table$Error))
        saveWorkbook(wb, file, overwrite = TRUE)
      } else {
        wb <- createWorkbook()
        
        addWorksheet(wb, "Life Table")
        writeData(wb, "Life Table", results$table)
        
        addWorksheet(wb, "Summary")
        summary_df <- data.frame(
          Parameter = c("Analysis Date", "Total Observations", "Total Events", 
                        "Total Censored", "Number of Time Points", "Time Range", "Time Points"),
          Value = c(
            as.character(Sys.Date()),
            nrow(surv_results()$data),
            sum(surv_results()$data$event),
            nrow(surv_results()$data) - sum(surv_results()$data$event),
            length(results$times),
            paste(min(results$times), "to", max(results$times)),
            paste(results$times, collapse = ", ")
          )
        )
        writeData(wb, "Summary", summary_df)
        
        addWorksheet(wb, "Parameters")
        params_df <- data.frame(
          Parameter = c("Time Variable", "Event Variable", "Grouping Variable", 
                        "Time Points", "Decimal Places"),
          Value = c(
            input$surv_time,
            input$surv_event,
            ifelse(input$surv_group != "none", input$surv_group, "None"),
            input$life_table_times,
            input$life_table_decimals
          )
        )
        writeData(wb, "Parameters", params_df)
        
        style <- createStyle(halign = "center", valign = "center")
        addStyle(wb, "Life Table", style, rows = 1:(nrow(results$table) + 1), 
                 cols = 1:ncol(results$table), gridExpand = TRUE)
        
        setColWidths(wb, "Life Table", cols = 1:ncol(results$table), widths = "auto")
        setColWidths(wb, "Summary", cols = 1:2, widths = "auto")
        setColWidths(wb, "Parameters", cols = 1:2, widths = "auto")
        
        saveWorkbook(wb, file, overwrite = TRUE)
      }
    }
  )
  
  # Auto-update life table when survival results change
  observeEvent(surv_results(), {
    if (!is.null(surv_results()) && surv_results()$success) {
      life_table_results()
    }
  })
  
  # Reactive for plot options to update plot without re-running analysis
  plot_options <- reactive({
    list(
      surv_ci = input$surv_ci,
      surv_risktable = input$surv_risktable,
      surv_median = input$surv_median,
      surv_labels = input$surv_labels,
      surv_color1 = input$surv_color1,
      surv_color2 = input$surv_color2,
      surv_color3 = input$surv_color3,
      surv_title = input$surv_title,
      surv_xlab = input$surv_xlab,
      surv_ylab = input$surv_ylab,
      surv_title_size = input$surv_title_size,
      surv_x_size = input$surv_x_size,
      surv_y_size = input$surv_y_size,
      surv_label_size = input$surv_label_size
    )
  })
  
  # Update the surv_plot output to use plot_options
  output$surv_plot <- renderPlot({
    req(surv_results(), plot_options())
    
    results <- surv_results()
    opts <- plot_options()
    
    if (!results$success) {
      p <- ggplot() +
        annotate("text", x = 0.5, y = 0.5, 
                 label = paste("Error:", results$error), 
                 size = 5, color = "red", hjust = 0.5, vjust = 0.5) +
        theme_void() +
        labs(title = "Kaplan-Meier Plot - Error")
      return(p)
    }
    
    tryCatch({
      if (results$n_groups == 1) {
        colors <- opts$surv_color1
      } else if (results$n_groups == 2) {
        colors <- c(opts$surv_color1, opts$surv_color2)
      } else {
        colors <- c(opts$surv_color1, opts$surv_color2, opts$surv_color3)
        if (results$n_groups > 3) {
          colors <- RColorBrewer::brewer.pal(results$n_groups, "Set1")
        }
      }
      
      # Ensure font sizes are valid numbers
      title_size <- ifelse(is.null(opts$surv_title_size) || opts$surv_title_size < 1, 
                           18, opts$surv_title_size)
      x_size <- ifelse(is.null(opts$surv_x_size) || opts$surv_x_size < 1, 
                       14, opts$surv_x_size)
      y_size <- ifelse(is.null(opts$surv_y_size) || opts$surv_y_size < 1, 
                       14, opts$surv_y_size)
      label_size <- ifelse(is.null(opts$surv_label_size) || opts$surv_label_size < 1, 
                           5, opts$surv_label_size)
      
      # Ensure title and labels are not NULL
      plot_title <- ifelse(is.null(opts$surv_title) || opts$surv_title == "", 
                           "Kaplan-Meier Curve", opts$surv_title)
      x_lab <- ifelse(is.null(opts$surv_xlab) || opts$surv_xlab == "", 
                      "Time", opts$surv_xlab)
      y_lab <- ifelse(is.null(opts$surv_ylab) || opts$surv_ylab == "", 
                      "Survival Probability", opts$surv_ylab)
      
      # Create Kaplan-Meier plot with safe parameters
      p <- ggsurvplot(
        results$fit,
        data = results$data,
        conf.int = opts$surv_ci,
        risk.table = opts$surv_risktable,
        palette = colors,
        title = plot_title,
        xlab = x_lab,
        ylab = y_lab,
        legend.title = ifelse(input$surv_group != "none", input$surv_group, ""),
        ggtheme = theme_minimal(),
        font.title = c(max(title_size, 1), "bold"),  # Ensure at least size 1
        font.x = c(max(x_size, 1), "plain"),         # Ensure at least size 1
        font.y = c(max(y_size, 1), "plain"),         # Ensure at least size 1
        font.tickslab = c(12, "plain"),
        risk.table.height = 0.25,
        surv.median.line = ifelse(opts$surv_median, "hv", "none"),
        censor = TRUE,
        size = 1.2,
        risk.table.fontsize = 4.5  # Explicitly set risk table font size
      )
      
      # Apply additional theme customizations
      p$plot <- p$plot +
        theme(
          plot.title = element_text(size = max(title_size, 1), face = "bold", hjust = 0.5),
          axis.title.x = element_text(size = max(x_size, 1), face = "bold"),
          axis.title.y = element_text(size = max(y_size, 1), face = "bold"),
          axis.text.x = element_text(size = max(x_size - 2, 10)),
          axis.text.y = element_text(size = max(y_size - 2, 10)),
          legend.title = element_text(size = max(x_size, 12), face = "bold"),
          legend.text = element_text(size = max(x_size - 2, 10)),
          legend.position = "right"
        )
      
      if (opts$surv_labels) {
        surv_summary <- summary(results$fit)
        
        if (results$n_groups == 1) {
          key_times <- quantile(surv_summary$time, probs = seq(0.1, 0.9, 0.2), na.rm = TRUE)
          key_points <- surv_summary$surv[findInterval(key_times, surv_summary$time)]
          
          # Ensure we have valid data points
          valid_indices <- !is.na(key_times) & !is.na(key_points)
          if (any(valid_indices)) {
            p$plot <- p$plot +
              geom_point(data = data.frame(time = key_times[valid_indices], 
                                           surv = key_points[valid_indices]), 
                         aes(x = time, y = surv), size = 3, color = "darkred") +
              geom_text(data = data.frame(time = key_times[valid_indices], 
                                          surv = key_points[valid_indices]),
                        aes(x = time, y = surv, 
                            label = paste0("S(t)=", round(surv, 2), "\nt=", round(time, 1))),
                        size = max(label_size, 3), vjust = -0.5, hjust = -0.1, color = "darkblue")
          }
        } else {
          if (!is.null(results$median_survival)) {
            median_data <- results$median_survival
            median_data$surv <- 0.5
            
            p$plot <- p$plot +
              geom_point(data = median_data, 
                         aes(x = median, y = surv, color = strata), 
                         size = 3, shape = 18) +
              geom_text(data = median_data,
                        aes(x = median, y = surv, color = strata,
                            label = paste0("Median: ", round(median, 1))),
                        size = max(label_size, 3), vjust = -1, hjust = -0.1, 
                        show.legend = FALSE)
          }
        }
      }
      
      if (opts$surv_risktable) {
        p$table <- p$table +
          theme(
            axis.title.x = element_text(size = max(x_size - 2, 10), face = "bold"),
            axis.text.x = element_text(size = max(x_size - 2, 10)),
            plot.title = element_text(size = max(title_size - 2, 12), face = "bold")
          )
      }
      
      return(p)
      
    }, error = function(e) {
      # Create a simple ggplot error plot as fallback
      ggplot() +
        annotate("text", x = 0.5, y = 0.5, 
                 label = paste("Plotting Error:", e$message), 
                 size = 5, color = "red", hjust = 0.5, vjust = 0.5) +
        theme_void() +
        labs(title = "Kaplan-Meier Plot - Error")
    })
  }, height = 600)
  
  # Add summary statistics UI
  output$surv_stats <- renderUI({
    req(surv_results())
    
    results <- surv_results()
    
    if (!results$success) {
      return(
        div(
          class = "alert alert-danger",
          h4("Kaplan-Meier Analysis Error"),
          p(results$error)
        )
      )
    }
    
    div(
      class = "card",
      div(
        class = "card-header",
        h4("Survival Analysis Summary", class = "card-title", style = "margin: 0;")
      ),
      div(
        class = "card-body",
        fluidRow(
          column(6,
                 tags$ul(
                   class = "list-unstyled",
                   tags$li(tags$strong("Total Observations:"), nrow(results$data)),
                   tags$li(tags$strong("Events:"), sum(results$data$event)),
                   tags$li(tags$strong("Censored:"), sum(results$data$event == 0))
                 )
          ),
          column(6,
                 tags$ul(
                   class = "list-unstyled",
                   tags$li(tags$strong("Groups:"), results$n_groups),
                   if (results$n_groups > 1 && !is.null(results$logrank_test)) {
                     tagList(
                       tags$li(tags$strong("Log-Rank p-value:"), 
                               round(1 - pchisq(results$logrank_test$chisq, length(results$logrank_test$n) - 1), 4))
                     )
                   }
                 )
          )
        )
      )
    )
  })
  
  output$surv_summary <- renderPrint({
    req(surv_results())
    
    results <- surv_results()
    
    if (!results$success) {
      cat("Kaplan-Meier Analysis Error:\n")
      cat(results$error, "\n")
      return()
    }
    
    cat("Kaplan-Meier Survival Analysis Summary\n")
    cat("======================================\n\n")
    
    print(results$fit)
    
    if (!is.null(results$median_survival)) {
      cat("\nMedian Survival Time:\n")
      print(results$median_survival)
    }
    
    if (!is.null(results$logrank_test)) {
      cat("\nLog-Rank Test (Group Comparison):\n")
      print(results$logrank_test)
      cat("p-value:", round(1 - pchisq(results$logrank_test$chisq, length(results$logrank_test$n) - 1), 4), "\n")
    }
    
    cat("\nSummary Statistics:\n")
    cat("- Number of observations:", nrow(results$data), "\n")
    cat("- Number of events:", sum(results$data$event), "\n")
    cat("- Censored:", sum(results$data$event == 0), "\n")
    
    if (results$n_groups > 1) {
      cat("- Number of groups:", results$n_groups, "\n")
      cat("- Group levels:", paste(results$group_levels, collapse = ", "), "\n")
    }
  })
  
  output$download_surv <- downloadHandler(
    filename = function() {
      paste("kaplan_meier_plot_", Sys.Date(), ".png", sep = "")
    },
    content = function(file) {
      req(surv_results(), plot_options())
      
      results <- surv_results()
      opts <- plot_options()
      
      if (!results$success) {
        p <- ggplot() +
          annotate("text", x = 0.5, y = 0.5, 
                   label = paste("Error:", results$error), 
                   size = 5, color = "red", hjust = 0.5, vjust = 0.5) +
          theme_void() +
          labs(title = "Kaplan-Meier Plot - Error")
        
        ggsave(file, plot = p, device = "png", width = 12, height = 10, dpi = 300)
      } else {
        if (results$n_groups == 1) {
          colors <- opts$surv_color1
        } else if (results$n_groups == 2) {
          colors <- c(opts$surv_color1, opts$surv_color2)
        } else {
          colors <- c(opts$surv_color1, opts$surv_color2, opts$surv_color3)
          if (results$n_groups > 3) {
            colors <- RColorBrewer::brewer.pal(results$n_groups, "Set1")
          }
        }
        
        p_download <- ggsurvplot(
          results$fit,
          data = results$data,
          conf.int = opts$surv_ci,
          risk.table = opts$surv_risktable,
          palette = colors,
          title = opts$surv_title,
          xlab = opts$surv_xlab,
          ylab = opts$surv_ylab,
          legend.title = ifelse(input$surv_group != "none", input$surv_group, ""),
          ggtheme = theme_minimal(),
          font.title = c(18, "bold"),
          font.x = c(14, "bold"),
          font.y = c(14, "bold"),
          font.tickslab = c(12, "plain"),
          risk.table.height = 0.25,
          surv.median.line = ifelse(opts$surv_median, "hv", "none"),
          size = 1.5
        )
        
        if (opts$surv_labels) {
          surv_summary <- summary(results$fit)
          
          if (results$n_groups == 1) {
            key_times <- quantile(surv_summary$time, probs = seq(0.1, 0.9, 0.2))
            key_points <- surv_summary$surv[findInterval(key_times, surv_summary$time)]
            
            p_download$plot <- p_download$plot +
              geom_point(data = data.frame(time = key_times, surv = key_points), 
                         aes(x = time, y = surv), size = 4, color = "darkred") +
              geom_text(data = data.frame(time = key_times, surv = key_points),
                        aes(x = time, y = surv, 
                            label = paste0("S(t)=", round(surv, 2), "\nt=", round(time, 1))),
                        size = 5, vjust = -0.5, hjust = -0.1, color = "darkblue")
          }
        }
        
        p_download$plot <- p_download$plot +
          theme(
            plot.title = element_text(size = 18, face = "bold", hjust = 0.5),
            axis.title.x = element_text(size = 14, face = "bold"),
            axis.title.y = element_text(size = 14, face = "bold"),
            legend.title = element_text(size = 12, face = "bold"),
            legend.text = element_text(size = 11)
          )
        
        if (opts$surv_risktable) {
          p_download$table <- p_download$table +
            theme(
              axis.title.x = element_text(size = 12, face = "bold"),
              axis.text.x = element_text(size = 11),
              plot.title = element_text(size = 12, face = "bold")
            )
        }
        
        ggsave(file, plot = print(p_download), device = "png", 
               width = 14, height = 10, dpi = 300, bg = "white")
      }
    }
  )
  
  # Add data download
  output$download_surv_data <- downloadHandler(
    filename = function() {
      paste("survival_data_", Sys.Date(), ".csv", sep = "")
    },
    content = function(file) {
      req(surv_results())
      
      results <- surv_results()
      
      if (!results$success) {
        write.csv(data.frame(Error = results$error), file, row.names = FALSE)
      } else {
        surv_data <- results$data %>%
          select(Time = time, Event = event)
        
        if (input$surv_group != "none") {
          surv_data$Group <- results$data$group
        }
        
        write.csv(surv_data, file, row.names = FALSE)
      }
    }
  )
  
  # Cox Regression reference UI
  output$cox_ref_ui <- renderUI({
    req(input$cox_vars, processed_data())
    df <- processed_data()
    
    ref_inputs <- tagList()
    
    for(var in input$cox_vars) {
      if(is.factor(df[[var]]) || is.character(df[[var]])) {
        levels <- if(is.factor(df[[var]])) levels(df[[var]]) else unique(na.omit(df[[var]]))
        ref_inputs <- tagAppendChild(
          ref_inputs,
          selectInput(paste0("cox_ref_", var), 
                      paste("Reference level for", var), 
                      choices = levels)
        )
      }
    }
    
    ref_inputs
  })
  
  # Cox Regression with action button
  cox_results <- eventReactive(input$run_cox, {
    req(input$cox_time, input$cox_event, input$cox_vars, processed_data())
    df <- processed_data()
    
    # Set reference levels
    for(var in input$cox_vars) {
      ref_input <- input[[paste0("cox_ref_", var)]]
      if(!is.null(ref_input) && (is.factor(df[[var]]) || is.character(df[[var]]))) {
        df[[var]] <- factor(df[[var]])
        df[[var]] <- relevel(df[[var]], ref = ref_input)
      }
    }
    
    # Create survival object
    surv_obj <- create_surv_obj(input$cox_time, input$cox_event, df)
    
    # Create formula
    formula <- as.formula(paste("surv_obj ~", paste(input$cox_vars, collapse = "+")))
    
    # Fit Cox model
    model <- coxph(formula, data = df)
    
    list(model = model, summary = summary(model))
  })
  
  output$cox_summary <- renderPrint({
    req(cox_results())
    print(cox_results()$summary)
  })
  
  output$cox_table <- renderDT({
    req(cox_results())
    cox_table <- broom::tidy(cox_results()$model, conf.int = TRUE, conf.level = input$cox_conf, exponentiate = TRUE)
    datatable(cox_table, options = list(pageLength = 10))
  })
  
  output$download_cox <- downloadHandler(
    filename = "cox_regression.docx",
    content = function(file) {
      cox_table <- broom::tidy(cox_results()$model, conf.int = TRUE, conf.level = input$cox_conf, exponentiate = TRUE)
      
      ft <- flextable(cox_table) %>%
        set_caption("Cox Regression Results") %>%
        theme_zebra()
      
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Poisson Regression reference UI
  output$poisson_ref_ui <- renderUI({
    req(input$poisson_vars, processed_data())
    df <- processed_data()
    
    ref_inputs <- tagList()
    
    for(var in input$poisson_vars) {
      if(is.factor(df[[var]]) || is.character(df[[var]])) {
        levels <- if(is.factor(df[[var]])) levels(df[[var]]) else unique(na.omit(df[[var]]))
        ref_inputs <- tagAppendChild(
          ref_inputs,
          selectInput(paste0("poisson_ref_", var), 
                      paste("Reference level for", var), 
                      choices = levels)
        )
      }
    }
    
    ref_inputs
  })
  
  # Poisson Regression with action button - CORRECTED VERSION
  poisson_results <- eventReactive(input$run_poisson, {
    req(input$poisson_outcome, input$poisson_vars, processed_data())
    
    tryCatch({
      df <- processed_data()
      
      # Validate variables exist
      if (!input$poisson_outcome %in% names(df)) {
        return(list(
          error = "Outcome variable not found in dataset",
          success = FALSE
        ))
      }
      
      missing_vars <- setdiff(input$poisson_vars, names(df))
      if (length(missing_vars) > 0) {
        return(list(
          error = paste("Predictor variables not found:", paste(missing_vars, collapse = ", ")),
          success = FALSE
        ))
      }
      
      # Set reference levels
      for(var in input$poisson_vars) {
        ref_input <- input[[paste0("poisson_ref_", var)]]
        if(!is.null(ref_input) && (is.factor(df[[var]]) || is.character(df[[var]]))) {
          df[[var]] <- factor(df[[var]])
          df[[var]] <- relevel(df[[var]], ref = ref_input)
        }
      }
      
      # Create formula
      formula <- as.formula(paste(input$poisson_outcome, "~", paste(input$poisson_vars, collapse = "+")))
      
      # Add offset if specified
      if(input$poisson_offset && !is.null(input$poisson_offset_var)) {
        if (!input$poisson_offset_var %in% names(df)) {
          return(list(
            error = "Offset variable not found in dataset",
            success = FALSE
          ))
        }
        formula <- as.formula(paste(input$poisson_outcome, "~", paste(input$poisson_vars, collapse = "+"), 
                                    "+ offset(log(", input$poisson_offset_var, "))"))
      }
      
      # Remove rows with missing values
      df_complete <- df[complete.cases(df[, c(input$poisson_outcome, input$poisson_vars)]), ]
      
      if (nrow(df_complete) == 0) {
        return(list(
          error = "No complete cases after removing missing values",
          success = FALSE
        ))
      }
      
      # Check if outcome is count data (non-negative integers)
      outcome_data <- df_complete[[input$poisson_outcome]]
      if (any(outcome_data < 0, na.rm = TRUE)) {
        return(list(
          error = "Outcome variable contains negative values. Poisson regression requires non-negative counts.",
          success = FALSE
        ))
      }
      
      # Fit Poisson model
      model <- glm(formula, family = poisson(link = "log"), data = df_complete)
      
      # Check for model convergence
      if (!model$converged) {
        warning("Poisson model did not converge properly")
      }
      
      # Calculate confidence intervals
      conf_level <- input$poisson_conf
      z_value <- qnorm(1 - (1 - conf_level) / 2)
      
      list(
        model = model,
        summary = summary(model),
        data = df_complete,
        success = TRUE
      )
      
    }, error = function(e) {
      return(list(
        error = paste("Error in Poisson regression:", e$message),
        success = FALSE
      ))
    })
  })
  
  output$poisson_summary <- renderPrint({
    req(poisson_results())
    
    results <- poisson_results()
    
    if (!results$success) {
      cat("POISSON REGRESSION ANALYSIS - ERROR\n")
      cat("===================================\n")
      cat(results$error, "\n")
      return()
    }
    
    cat("POISSON REGRESSION ANALYSIS\n")
    cat("===========================\n\n")
    
    # Model information
    cat("MODEL INFORMATION:\n")
    cat("- Outcome variable:", input$poisson_outcome, "\n")
    cat("- Predictor variables:", paste(input$poisson_vars, collapse = ", "), "\n")
    if (input$poisson_offset && !is.null(input$poisson_offset_var)) {
      cat("- Offset variable:", input$poisson_offset_var, "\n")
    }
    cat("- Observations:", nrow(results$data), "\n")
    cat("- Confidence level:", input$poisson_conf * 100, "%\n")
    cat("\n")
    
    # Model summary
    print(results$summary)
    
    # Additional model diagnostics
    cat("\nMODEL DIAGNOSTICS:\n")
    cat("- Null deviance:", round(results$model$null.deviance, 3), "\n")
    cat("- Residual deviance:", round(results$model$deviance, 3), "\n")
    cat("- AIC:", round(AIC(results$model), 3), "\n")
    cat("- Dispersion parameter:", round(results$summary$dispersion, 3), "\n")
    
    # Check for overdispersion
    if (results$summary$dispersion > 1.5) {
      cat("⚠ Warning: Possible overdispersion detected. Consider using negative binomial regression.\n")
    }
  })
  
  output$poisson_table <- renderDT({
    req(poisson_results())
    
    results <- poisson_results()
    
    if (!results$success) {
      return(datatable(
        data.frame(Error = results$error),
        options = list(dom = 't'),
        rownames = FALSE,
        caption = "Poisson Regression - Error"
      ))
    }
    
    # Create results table with rate ratios and confidence intervals
    poisson_table <- broom::tidy(results$model, conf.int = TRUE, conf.level = input$poisson_conf, exponentiate = TRUE)
    
    # Add log-scale estimates
    poisson_table_log <- broom::tidy(results$model, conf.int = TRUE, conf.level = input$poisson_conf, exponentiate = FALSE)
    poisson_table$log_estimate <- poisson_table_log$estimate
    poisson_table$log_std.error <- poisson_table_log$std.error
    poisson_table$log_conf.low <- poisson_table_log$conf.low
    poisson_table$log_conf.high <- poisson_table_log$conf.high
    
    datatable(
      poisson_table,
      options = list(
        pageLength = 15,
        dom = 'Blfrtip',
        scrollX = TRUE,
        buttons = c('copy', 'csv', 'excel')
      ),
      rownames = FALSE,
      caption = "Poisson Regression Results (Rate Ratios and 95% Confidence Intervals)",
      extensions = 'Buttons'
    ) %>%
      formatRound(columns = c('estimate', 'std.error', 'statistic', 'p.value', 'conf.low', 'conf.high',
                              'log_estimate', 'log_std.error', 'log_conf.low', 'log_conf.high'), 
                  digits = 4)
  })
  
  output$download_poisson <- downloadHandler(
    filename = function() {
      paste("poisson_regression_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(poisson_results())
      
      results <- poisson_results()
      
      if (!results$success) {
        ft <- flextable(data.frame(Error = results$error)) %>%
          set_caption("Poisson Regression Analysis - Error") %>%
          theme_zebra()
        
        doc <- read_docx() %>% 
          body_add_flextable(ft)
      } else {
        poisson_table <- broom::tidy(results$model, conf.int = TRUE, conf.level = input$poisson_conf, exponentiate = TRUE)
        
        poisson_table_log <- broom::tidy(results$model, conf.int = TRUE, conf.level = input$poisson_conf, exponentiate = FALSE)
        poisson_table$log_estimate <- poisson_table_log$estimate
        poisson_table$log_std.error <- poisson_table_log$std.error
        poisson_table$log_conf.low <- poisson_table_log$conf.low
        poisson_table$log_conf.high <- poisson_table_log$conf.high
        
        doc <- read_docx()
        
        doc <- doc %>%
          body_add_par("POISSON REGRESSION ANALYSIS REPORT", style = "heading 1") %>%
          body_add_par(paste("Outcome Variable:", input$poisson_outcome), style = "Normal") %>%
          body_add_par(paste("Predictor Variables:", paste(input$poisson_vars, collapse = ", ")), style = "Normal")
        
        if (input$poisson_offset && !is.null(input$poisson_offset_var)) {
          doc <- doc %>% body_add_par(paste("Offset Variable:", input$poisson_offset_var), style = "Normal")
        }
        
        doc <- doc %>%
          body_add_par(paste("Confidence Level:", input$poisson_conf * 100, "%"), style = "Normal") %>%
          body_add_par(paste("Observations:", nrow(results$data)), style = "Normal") %>%
          body_add_par("", style = "Normal")
        
        doc <- doc %>%
          body_add_par("REGRESSION COEFFICIENTS (RATE RATIOS)", style = "heading 2")
        
        ft <- flextable(poisson_table) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft)
        
        doc <- doc %>%
          body_add_par("MODEL DIAGNOSTICS", style = "heading 2")
        
        diagnostics <- data.frame(
          Statistic = c("Null Deviance", "Residual Deviance", "AIC", "Dispersion Parameter"),
          Value = c(
            round(results$model$null.deviance, 3),
            round(results$model$deviance, 3),
            round(AIC(results$model), 3),
            round(results$summary$dispersion, 3)
          )
        )
        
        ft_diag <- flextable(diagnostics) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_diag)
      }
      
      print(doc, target = file)
    }
  )
  
  # Linear Regression reference UI
  output$linear_ref_ui <- renderUI({
    req(input$linear_vars, processed_data())
    df <- processed_data()
    
    ref_inputs <- tagList()
    
    for(var in input$linear_vars) {
      if(is.factor(df[[var]]) || is.character(df[[var]])) {
        levels <- if(is.factor(df[[var]])) levels(df[[var]]) else unique(na.omit(df[[var]]))
        ref_inputs <- tagAppendChild(
          ref_inputs,
          selectInput(paste0("linear_ref_", var), 
                      paste("Reference level for", var), 
                      choices = levels)
        )
      }
    }
    
    ref_inputs
  })
  
  # Linear Regression with action button
  linear_results <- eventReactive(input$run_linear, {
    req(input$linear_outcome, input$linear_vars, processed_data())
    df <- processed_data()
    
    # Set reference levels
    for(var in input$linear_vars) {
      ref_input <- input[[paste0("linear_ref_", var)]]
      if(!is.null(ref_input) && (is.factor(df[[var]]) || is.character(df[[var]]))) {
        df[[var]] <- factor(df[[var]])
        df[[var]] <- relevel(df[[var]], ref = ref_input)
      }
    }
    
    # Create formula
    formula <- as.formula(paste(input$linear_outcome, "~", paste(input$linear_vars, collapse = "+")))
    
    # Fit linear model
    model <- lm(formula, data = df)
    
    list(model = model, summary = summary(model))
  })
  
  output$linear_summary <- renderPrint({
    req(linear_results())
    print(linear_results()$summary)
    
    if(input$linear_std) {
      cat("\nStandardized Coefficients:\n")
      std_coef <- lm.beta(linear_results()$model)
      print(std_coef)
    }
  })
  
  output$linear_table <- renderDT({
    req(linear_results())
    linear_table <- broom::tidy(linear_results()$model, conf.int = TRUE)
    
    if(input$linear_std) {
      std_coef <- lm.beta(linear_results()$model)$standardized.coefficients
      linear_table$std_estimate <- std_coef
    }
    
    datatable(linear_table, options = list(pageLength = 10))
  })
  
  output$download_linear <- downloadHandler(
    filename = "linear_regression.docx",
    content = function(file) {
      linear_table <- broom::tidy(linear_results()$model, conf.int = TRUE)
      
      ft <- flextable(linear_table) %>%
        set_caption("Linear Regression Results") %>%
        theme_zebra()
      
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  # Server logic for logistic regression - CORRECTED VERSION (3 decimal places)
  
  # UI for outcome variable handling
  output$logistic_outcome_ui <- renderUI({
    req(input$logistic_outcome, processed_data())
    df <- processed_data()
    
    if (!input$logistic_outcome %in% names(df)) return(NULL)
    
    outcome_var <- df[[input$logistic_outcome]]
    
    # Check if outcome is numeric binary (0/1) or needs conversion
    if (is.numeric(outcome_var)) {
      unique_vals <- unique(na.omit(outcome_var))
      if (length(unique_vals) == 2 && all(unique_vals %in% c(0, 1))) {
        return(
          div(
            class = "alert alert-success",
            h5("Outcome Variable Status"),
            p("✓ Outcome variable is numeric binary (0/1)"),
            p(paste("Levels: 0 =", names(which.min(table(outcome_var))), 
                    ", 1 =", names(which.max(table(outcome_var)))))
          )
        )
      }
    }
    
    # If not numeric binary, provide reference level selection
    tagList(
      div(
        class = "alert alert-info",
        h5("Outcome Variable Configuration"),
        p("Please specify the reference category for the outcome variable:")
      ),
      selectInput("logistic_outcome_ref", 
                  "Reference Category (Baseline):",
                  choices = if (is.factor(outcome_var)) {
                    levels(outcome_var)
                  } else {
                    unique(na.omit(outcome_var))
                  }),
      helpText("The reference category will be coded as 0 in the analysis.")
    )
  })
  
  # UI for predictor variable reference levels
  output$logistic_ref_ui <- renderUI({
    req(input$logistic_vars, processed_data())
    df <- processed_data()
    
    ref_inputs <- tagList(
      h5("Reference Levels for Predictor Variables"),
      p("Specify reference categories for categorical predictors:")
    )
    
    for(var in input$logistic_vars) {
      if (var %in% names(df)) {
        var_data <- df[[var]]
        
        # Only show reference selection for categorical variables with reasonable number of levels
        if ((is.factor(var_data) || is.character(var_data)) && length(unique(na.omit(var_data))) <= 20) {
          
          levels <- if (is.factor(var_data)) {
            levels(var_data)
          } else {
            sort(unique(na.omit(var_data)))
          }
          
          ref_inputs <- tagAppendChild(
            ref_inputs,
            selectInput(paste0("logistic_ref_", make.names(var)), 
                        paste("Reference for", var), 
                        choices = levels,
                        selected = levels[1])
          )
        } else if ((is.factor(var_data) || is.character(var_data)) && length(unique(na.omit(var_data))) > 20) {
          ref_inputs <- tagAppendChild(
            ref_inputs,
            div(
              class = "alert alert-warning",
              p(paste("⚠ Variable", var, "has", length(unique(na.omit(var_data))), 
                      "levels. Consider recoding or using as continuous variable."))
            )
          )
        }
      }
    }
    
    if (length(ref_inputs) == 2) { # Only has header and no variables
      ref_inputs <- tagList(
        h5("Reference Levels for Predictor Variables"),
        p("No categorical predictors selected or all predictors are numeric.")
      )
    }
    
    ref_inputs
  })
  
  # Logistic Regression Analysis - CORRECTED VERSION (3 decimal places)
  logistic_results <- eventReactive(input$run_logistic, {
    req(input$logistic_outcome, input$logistic_vars, processed_data())
    
    tryCatch({
      df <- processed_data()
      
      # Validate variables exist
      if (!input$logistic_outcome %in% names(df)) {
        stop("Outcome variable not found in dataset")
      }
      
      missing_vars <- setdiff(input$logistic_vars, names(df))
      if (length(missing_vars) > 0) {
        stop(paste("Predictor variables not found:", paste(missing_vars, collapse = ", ")))
      }
      
      # Prepare outcome variable
      outcome_var <- df[[input$logistic_outcome]]
      
      # Convert outcome to binary numeric (0/1)
      if (is.numeric(outcome_var)) {
        unique_vals <- unique(na.omit(outcome_var))
        if (length(unique_vals) == 2 && all(unique_vals %in% c(0, 1))) {
          # Already numeric binary
          df$outcome_binary <- outcome_var
          outcome_levels <- c("0" = "Reference", "1" = "Comparison")
        } else {
          # Numeric but not binary - convert to binary based on median
          cutoff <- median(outcome_var, na.rm = TRUE)
          df$outcome_binary <- as.numeric(outcome_var > cutoff)
          outcome_levels <- c("0" = paste("≤", round(cutoff, 3)), "1" = paste(">", round(cutoff, 3)))
          warning("Numeric outcome converted to binary using median cutoff")
        }
      } else {
        # Categorical outcome - use user-specified reference
        if (is.null(input$logistic_outcome_ref)) {
          stop("Please specify reference category for outcome variable")
        }
        
        outcome_factor <- factor(outcome_var)
        ref_level <- input$logistic_outcome_ref
        
        # Set reference level
        if (ref_level %in% levels(outcome_factor)) {
          outcome_factor <- relevel(outcome_factor, ref = ref_level)
        }
        
        # Convert to binary numeric
        df$outcome_binary <- as.numeric(outcome_factor) - 1
        outcome_levels <- setNames(levels(outcome_factor), c("0", "1"))
      }
      
      # Remove rows with missing outcome
      df <- df[!is.na(df$outcome_binary), ]
      
      # Check if we have both outcome levels
      if (length(unique(df$outcome_binary)) != 2) {
        stop("Outcome variable must have exactly two categories after processing")
      }
      
      # Prepare predictor variables
      for(var in input$logistic_vars) {
        if (var %in% names(df)) {
          var_data <- df[[var]]
          
          # Handle categorical predictors
          if (is.factor(var_data) || is.character(var_data)) {
            # Get user-specified reference level
            ref_input_name <- paste0("logistic_ref_", make.names(var))
            ref_level <- input[[ref_input_name]]
            
            if (!is.null(ref_level) && ref_level %in% unique(na.omit(var_data))) {
              df[[var]] <- factor(var_data)
              df[[var]] <- relevel(df[[var]], ref = ref_level)
            } else {
              # Use default reference (first level)
              df[[var]] <- factor(var_data)
            }
          }
          
          # Handle numeric predictors - ensure they're numeric
          if (is.character(var_data) && !all(is.na(suppressWarnings(as.numeric(na.omit(var_data)))))) {
            df[[var]] <- as.numeric(var_data)
          }
        }
      }
      
      # Remove rows with missing values in predictors
      complete_cases <- complete.cases(df[, input$logistic_vars, drop = FALSE])
      df <- df[complete_cases, ]
      
      if (nrow(df) == 0) {
        stop("No complete cases after removing missing values")
      }
      
      # Check for complete separation or other issues
      if (any(table(df$outcome_binary) < 2)) {
        stop("Insufficient data in one or more outcome categories for logistic regression")
      }
      
      # Create formula
      formula <- as.formula(paste("outcome_binary ~", paste(input$logistic_vars, collapse = "+")))
      
      # Fit logistic regression model
      model <- glm(formula, family = binomial(link = "logit"), data = df)
      
      # Check model convergence
      if (!model$converged) {
        warning("Logistic regression model did not converge properly")
      }
      
      # Calculate odds ratios if requested
      conf_level <- 0.95  # Could make this user-configurable
      
      list(
        model = model,
        summary = summary(model),
        data = df,
        outcome_levels = outcome_levels,
        formula = formula,
        success = TRUE
      )
      
    }, error = function(e) {
      return(list(
        error = paste("Error in logistic regression:", e$message),
        success = FALSE
      ))
    })
  })
  
  output$logistic_summary <- renderPrint({
    req(logistic_results())
    
    results <- logistic_results()
    
    if (!results$success) {
      cat("LOGISTIC REGRESSION ANALYSIS - ERROR\n")
      cat("====================================\n")
      cat(results$error, "\n")
      return()
    }
    
    cat("LOGISTIC REGRESSION ANALYSIS\n")
    cat("============================\n\n")
    
    # Model information
    cat("MODEL INFORMATION:\n")
    cat("- Outcome variable:", input$logistic_outcome, "\n")
    cat("- Outcome levels:\n")
    for (level in names(results$outcome_levels)) {
      cat("  ", level, "=", results$outcome_levels[[level]], "\n")
    }
    cat("- Predictor variables:", paste(input$logistic_vars, collapse = ", "), "\n")
    cat("- Observations:", nrow(results$data), "\n")
    cat("- Model formula:", deparse(results$formula), "\n")
    cat("\n")
    
    # Model summary
    cat("MODEL SUMMARY:\n")
    print(results$summary)
    
    # Model diagnostics
    cat("\nMODEL DIAGNOSTICS:\n")
    cat("- Null deviance:", round(results$model$null.deviance, 3), "\n")
    cat("- Residual deviance:", round(results$model$deviance, 3), "\n")
    cat("- AIC:", round(AIC(results$model), 3), "\n")
    
    # Pseudo R-squared
    pseudo_r2 <- 1 - (results$model$deviance / results$model$null.deviance)
    cat("- McFadden's R²:", round(pseudo_r2, 3), "\n")
    
    # Check for multicollinearity
    if (length(input$logistic_vars) > 1) {
      vif_values <- tryCatch({
        car::vif(results$model)
      }, error = function(e) NULL)
      
      if (!is.null(vif_values)) {
        cat("- Variance Inflation Factors (VIF):\n")
        if (is.matrix(vif_values)) {
          vif_df <- as.data.frame(vif_values)
          colnames(vif_df) <- c("GVIF", "DF", "GVIF^(1/(2*DF))")
          print(vif_df)
        } else {
          print(round(vif_values, 3))
        }
        
        # Check for high VIF
        high_vif <- if (is.matrix(vif_values)) {
          vif_values[, "GVIF^(1/(2*DF))"] > 5
        } else {
          vif_values > 10
        }
        if (any(high_vif, na.rm = TRUE)) {
          cat("⚠ Warning: High multicollinearity detected (VIF > 10 or GVIF^(1/(2*DF)) > 5)\n")
        }
      }
    }
  })
  
  output$logistic_table <- renderDT({
    req(logistic_results())
    
    results <- logistic_results()
    
    if (!results$success) {
      return(datatable(
        data.frame(Error = results$error),
        options = list(dom = 't'),
        rownames = FALSE,
        caption = "Logistic Regression - Error"
      ))
    }
    
    # Create results table
    if (input$logistic_or) {
      # Show odds ratios
      logistic_table <- broom::tidy(results$model, conf.int = TRUE, conf.level = 0.95, exponentiate = TRUE)
    } else {
      # Show coefficients
      logistic_table <- broom::tidy(results$model, conf.int = TRUE, conf.level = 0.95, exponentiate = FALSE)
    }
    
    # Format for display
    logistic_table$p.value <- ifelse(logistic_table$p.value < 0.001, "<0.001", 
                                     round(logistic_table$p.value, 3))
    
    # Round all numeric columns to 3 decimal places
    numeric_cols <- sapply(logistic_table, is.numeric)
    logistic_table[numeric_cols] <- lapply(logistic_table[numeric_cols], function(x) {
      ifelse(is.na(x), NA, round(x, 3))
    })
    
    datatable(
      logistic_table,
      options = list(
        pageLength = 15,
        dom = 'Blfrtip',
        scrollX = TRUE,
        buttons = c('copy', 'csv', 'excel')
      ),
      rownames = FALSE,
      caption = ifelse(input$logistic_or,
                       "Logistic Regression Results (Odds Ratios and 95% Confidence Intervals)",
                       "Logistic Regression Results (Coefficients and 95% Confidence Intervals)"),
      extensions = 'Buttons'
    ) %>%
      formatStyle(
        'p.value',
        backgroundColor = styleInterval(
          c(0.001, 0.01, 0.05, 0.1),
          c('#D4EDDA', '#D1ECF1', '#FFF3CD', '#F8D7DA', '#FFFFFF')
        )
      )
  })
  
  output$download_logistic <- downloadHandler(
    filename = function() {
      paste("logistic_regression_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(logistic_results())
      
      results <- logistic_results()
      
      if (!results$success) {
        ft <- flextable(data.frame(Error = results$error)) %>%
          set_caption("Logistic Regression Analysis - Error") %>%
          theme_zebra()
        
        doc <- read_docx() %>% 
          body_add_flextable(ft)
      } else {
        # Create both coefficient tables
        coeff_table <- broom::tidy(results$model, conf.int = TRUE, conf.level = 0.95, exponentiate = FALSE)
        or_table <- broom::tidy(results$model, conf.int = TRUE, conf.level = 0.95, exponentiate = TRUE)
        
        # Round to 3 decimal places
        numeric_cols <- sapply(coeff_table, is.numeric)
        coeff_table[numeric_cols] <- lapply(coeff_table[numeric_cols], function(x) round(x, 3))
        or_table[numeric_cols] <- lapply(or_table[numeric_cols], function(x) round(x, 3))
        
        doc <- read_docx()
        
        doc <- doc %>%
          body_add_par("LOGISTIC REGRESSION ANALYSIS REPORT", style = "heading 1") %>%
          body_add_par(paste("Outcome Variable:", input$logistic_outcome), style = "Normal") %>%
          body_add_par("Outcome Variable Coding:", style = "Normal")
        
        # Add outcome level information
        for (level in names(results$outcome_levels)) {
          doc <- doc %>% body_add_par(paste("  ", level, "=", results$outcome_levels[[level]]), style = "Normal")
        }
        
        doc <- doc %>%
          body_add_par(paste("Predictor Variables:", paste(input$logistic_vars, collapse = ", ")), style = "Normal") %>%
          body_add_par(paste("Observations:", nrow(results$data)), style = "Normal") %>%
          body_add_par(paste("Model formula:", deparse(results$formula)), style = "Normal") %>%
          body_add_par("", style = "Normal")
        
        doc <- doc %>%
          body_add_par("REGRESSION COEFFICIENTS", style = "heading 2")
        
        ft_coeff <- flextable(coeff_table) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_coeff)
        
        doc <- doc %>%
          body_add_par("ODDS RATIOS", style = "heading 2")
        
        ft_or <- flextable(or_table) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_or)
        
        doc <- doc %>%
          body_add_par("MODEL DIAGNOSTICS", style = "heading 2")
        
        pseudo_r2 <- 1 - (results$model$deviance / results$model$null.deviance)
        
        diagnostics <- data.frame(
          Statistic = c("Null Deviance", "Residual Deviance", "AIC", "McFadden's R²"),
          Value = c(
            round(results$model$null.deviance, 3),
            round(results$model$deviance, 3),
            round(AIC(results$model), 3),
            round(pseudo_r2, 3)
          )
        )
        
        ft_diag <- flextable(diagnostics) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_diag)
        
        doc <- doc %>%
          body_add_par("INTERPRETATION GUIDE", style = "heading 2") %>%
          body_add_par("• Odds Ratio > 1: Increased odds of outcome", style = "Normal") %>%
          body_add_par("• Odds Ratio < 1: Decreased odds of outcome", style = "Normal") %>%
          body_add_par("• Odds Ratio = 1: No effect", style = "Normal") %>%
          body_add_par("• Confidence Interval that includes 1: Not statistically significant", style = "Normal") %>%
          body_add_par("• P-value < 0.05: Statistically significant", style = "Normal")
      }
      
      print(doc, target = file)
    }
  )
  
  # Chi-square Test with action button
  chisq_results <- eventReactive(input$run_chisq, {
    req(input$chisq_var1, input$chisq_var2, processed_data())
    df <- processed_data()
    
    # Create contingency table
    tab <- table(df[[input$chisq_var1]], df[[input$chisq_var2]])
    
    # Perform chi-square test
    chisq_test <- chisq.test(tab)
    
    list(
      table = tab,
      test = chisq_test,
      expected = if(input$chisq_expected) chisq_test$expected else NULL,
      residuals = if(input$chisq_residuals) chisq_test$residuals else NULL
    )
  })
  
  output$chisq_results <- renderPrint({
    req(chisq_results())
    
    cat("Chi-Square Test Results\n")
    cat("=======================\n\n")
    
    print(chisq_results()$test)
    
    if(input$chisq_fisher) {
      cat("\nFisher's Exact Test:\n")
      fisher_test <- fisher.test(chisq_results()$table)
      print(fisher_test)
    }
    
    if(input$chisq_rr) {
      cat("\nRelative Risk (RR) Calculation:\n")
      # Simple RR calculation
      tab <- chisq_results()$table
      if(nrow(tab) == 2 && ncol(tab) == 2) {
        rr <- (tab[2,2] / sum(tab[2,])) / (tab[1,2] / sum(tab[1,]))
        cat("Relative Risk:", round(rr, 3), "\n")
      }
    }
  })
  
  output$chisq_table <- renderDT({
    req(chisq_results())
    
    tab <- as.data.frame.matrix(chisq_results()$table)
    
    if(input$chisq_expected && !is.null(chisq_results()$expected)) {
      expected <- as.data.frame.matrix(chisq_results()$expected)
      colnames(expected) <- paste("Expected", colnames(expected))
      tab <- cbind(tab, expected)
    }
    
    if(input$chisq_residuals && !is.null(chisq_results()$residuals)) {
      residuals <- as.data.frame.matrix(chisq_results()$residuals)
      colnames(residuals) <- paste("Residual", colnames(residuals))
      tab <- cbind(tab, residuals)
    }
    
    datatable(tab, 
              options = list(pageLength = 10, scrollX = TRUE),
              caption = "Contingency Table")
  })
  
  output$chisq_plot <- renderPlot({
    req(chisq_results())
    
    tab <- chisq_results()$table
    
    # Create mosaic plot
    mosaicplot(tab, 
               main = "Mosaic Plot",
               xlab = input$chisq_var1,
               ylab = input$chisq_var2,
               color = TRUE,
               shade = TRUE)
  })
  
  output$download_chisq <- downloadHandler(
    filename = "chi_square_test.docx",
    content = function(file) {
      ft <- flextable(as.data.frame.matrix(chisq_results()$table)) %>%
        set_caption("Contingency Table") %>%
        theme_zebra()
      
      doc <- read_docx() %>% 
        body_add_flextable(ft)
      print(doc, target = file)
    }
  )
  
  output$download_chisq_plot <- downloadHandler(
    filename = "chi_square_plot.png",
    content = function(file) {
      png(file, width = 800, height = 600)
      mosaicplot(chisq_results()$table,
                 main = "Mosaic Plot",
                 xlab = input$chisq_var1,
                 ylab = input$chisq_var2,
                 color = TRUE,
                 shade = TRUE)
      dev.off()
    }
  )
  
  # ANOVA Analysis with action button
  anova_results <- eventReactive(input$run_anova, {
    req(input$anova_outcome, input$anova_group, processed_data())
    df <- processed_data()
    
    # Create formula
    formula <- as.formula(paste(input$anova_outcome, "~", input$anova_group))
    
    # Fit ANOVA model
    model <- aov(formula, data = df)
    
    # Post-hoc tests if requested
    posthoc <- NULL
    if(input$anova_posthoc) {
      if(input$anova_posthoc_method == "tukey") {
        posthoc <- TukeyHSD(model)
      } else if(input$anova_posthoc_method == "bonferroni") {
        posthoc <- pairwise.t.test(df[[input$anova_outcome]], 
                                   df[[input$anova_group]], 
                                   p.adjust.method = "bonferroni")
      }
    }
    
    # Assumptions check if requested
    assumptions <- NULL
    if(input$anova_assumptions) {
      # Normality test
      normality <- shapiro.test(residuals(model))
      
      # Homogeneity of variance
      homogeneity <- car::leveneTest(formula, data = df)
      
      assumptions <- list(
        normality = normality,
        homogeneity = homogeneity
      )
    }
    
    list(
      model = model,
      summary = summary(model),
      posthoc = posthoc,
      assumptions = assumptions
    )
  })
  
  output$anova_results <- renderPrint({
    req(anova_results())
    
    cat("One-Way ANOVA Results\n")
    cat("=====================\n\n")
    
    print(anova_results()$summary)
    
    if(!is.null(anova_results()$posthoc)) {
      cat("\nPost-hoc Comparisons:\n")
      print(anova_results()$posthoc)
    }
  })
  
  output$anova_table <- renderDT({
    req(anova_results())
    
    anova_table <- broom::tidy(anova_results()$model)
    datatable(anova_table, options = list(pageLength = 10))
  })
  
  output$anova_assumptions_results <- renderPrint({
    req(anova_results())
    
    if(!is.null(anova_results()$assumptions)) {
      cat("ANOVA Assumptions Check\n")
      cat("=======================\n\n")
      
      cat("1. Normality of Residuals (Shapiro-Wilk test):\n")
      print(anova_results()$assumptions$normality)
      
      cat("\n2. Homogeneity of Variance (Levene's test):\n")
      print(anova_results()$assumptions$homogeneity)
    }
  })
  
  output$anova_assumptions_plots <- renderPlot({
    req(anova_results())
    
    par(mfrow = c(2, 2))
    plot(anova_results()$model)
  })
  
  output$anova_posthoc_table <- renderDT({
    req(anova_results())
    
    if(!is.null(anova_results()$posthoc)) {
      if(input$anova_posthoc_method == "tukey") {
        posthoc_table <- as.data.frame(anova_results()$posthoc[[1]])
        datatable(posthoc_table, options = list(pageLength = 10))
      }
    }
  })
  
  output$anova_posthoc_plot <- renderPlot({
    req(anova_results())
    
    if(!is.null(anova_results()$posthoc) && input$anova_posthoc_method == "tukey") {
      plot(anova_results()$posthoc)
    }
  })
  
  # Two-Way ANOVA
  anova_2way_results <- eventReactive(input$run_anova_2way, {
    req(input$anova_outcome_2way, input$anova_factor1, input$anova_factor2, processed_data())
    df <- processed_data()
    
    # Create formula
    if(input$anova_interaction) {
      formula <- as.formula(paste(input$anova_outcome_2way, "~", 
                                  input$anova_factor1, "*", input$anova_factor2))
    } else {
      formula <- as.formula(paste(input$anova_outcome_2way, "~", 
                                  input$anova_factor1, "+", input$anova_factor2))
    }
    
    # Fit ANOVA model
    model <- aov(formula, data = df)
    
    # Simple effects analysis if requested
    simple_effects <- NULL
    if(input$anova_2way_posthoc && input$anova_interaction) {
      simple_effects <- emmeans(model, pairwise ~ input$anova_factor1 | input$anova_factor2)
    }
    
    list(
      model = model,
      summary = summary(model),
      simple_effects = simple_effects
    )
  })
  
  output$anova_2way_results <- renderPrint({
    req(anova_2way_results())
    
    cat("Two-Way ANOVA Results\n")
    cat("=====================\n\n")
    
    print(anova_2way_results()$summary)
  })
  
  output$anova_2way_table <- renderDT({
    req(anova_2way_results())
    
    anova_table <- broom::tidy(anova_2way_results()$model)
    datatable(anova_table, options = list(pageLength = 10))
  })
  
  output$anova_2way_posthoc <- renderPrint({
    req(anova_2way_results())
    
    if(!is.null(anova_2way_results()$simple_effects)) {
      cat("Simple Effects Analysis\n")
      cat("=======================\n\n")
      print(anova_2way_results()$simple_effects)
    }
  })
  
  # Welch ANOVA - CORRECTED VERSION
  welch_anova_results <- eventReactive(input$run_welch_anova, {
    req(input$welch_outcome, input$welch_group, processed_data())
    
    tryCatch({
      df <- processed_data()
      
      # Validate variables exist
      if (!input$welch_outcome %in% names(df)) {
        return(list(
          error = "Outcome variable not found in dataset",
          success = FALSE
        ))
      }
      
      if (!input$welch_group %in% names(df)) {
        return(list(
          error = "Grouping variable not found in dataset",
          success = FALSE
        ))
      }
      
      # Get the variables
      outcome_var <- df[[input$welch_outcome]]
      group_var <- df[[input$welch_group]]
      
      # Remove missing values
      complete_cases <- !is.na(outcome_var) & !is.na(group_var)
      outcome_clean <- outcome_var[complete_cases]
      group_clean <- group_var[complete_cases]
      
      if (length(outcome_clean) == 0) {
        return(list(
          error = "No complete cases after removing missing values",
          success = FALSE
        ))
      }
      
      # Convert to factor for grouping
      group_factor <- as.factor(group_clean)
      n_groups <- length(levels(group_factor))
      
      # Check for sufficient data
      group_counts <- table(group_factor)
      
      if (n_groups < 2) {
        return(list(
          error = "Grouping variable must have at least 2 categories",
          success = FALSE
        ))
      }
      
      # Check minimum group size (at least 2 observations per group for Welch ANOVA)
      if (any(group_counts < 2)) {
        small_groups <- names(group_counts)[group_counts < 2]
        return(list(
          error = paste("Groups with insufficient data (need at least 2 observations):", 
                        paste(small_groups, collapse = ", ")),
          success = FALSE
        ))
      }
      
      # Check variance within groups (avoid division by zero in Welch formula)
      group_variances <- tapply(outcome_clean, group_factor, var, na.rm = TRUE)
      if (any(group_variances == 0, na.rm = TRUE)) {
        zero_var_groups <- names(group_variances)[group_variances == 0]
        return(list(
          error = paste("Groups with zero variance (all values identical):", 
                        paste(zero_var_groups, collapse = ", ")),
          success = FALSE
        ))
      }
      
      # Perform Welch ANOVA
      welch_test <- oneway.test(outcome_clean ~ group_factor, 
                                data = data.frame(outcome_clean, group_factor),
                                var.equal = FALSE)
      
      # Check for NaN results
      if (is.nan(welch_test$statistic) || is.na(welch_test$p.value)) {
        return(list(
          error = "Welch ANOVA calculation failed (NaN/NA results). Possible causes: insufficient data, extreme outliers, or computational issues.",
          success = FALSE
        ))
      }
      
      # Games-Howell post-hoc test (if more than 2 groups)
      posthoc <- NULL
      if (n_groups > 2 && !is.nan(welch_test$statistic) && welch_test$p.value < 1) {
        posthoc <- tryCatch({
          # Using pairwise.t.test with Holm correction as alternative to Games-Howell
          pairwise.t.test(outcome_clean, group_factor, 
                          pool.sd = FALSE,
                          p.adjust.method = "holm")
        }, error = function(e) {
          NULL
        })
      }
      
      # Calculate descriptive statistics
      group_stats <- data.frame(
        Group = levels(group_factor),
        N = as.numeric(group_counts),
        Mean = tapply(outcome_clean, group_factor, mean, na.rm = TRUE),
        SD = tapply(outcome_clean, group_factor, sd, na.rm = TRUE),
        SE = tapply(outcome_clean, group_factor, function(x) sd(x, na.rm = TRUE)/sqrt(length(x))),
        Variance = group_variances
      )
      
      list(
        welch_test = welch_test,
        posthoc = posthoc,
        group_stats = group_stats,
        n_groups = n_groups,
        total_n = length(outcome_clean),
        success = TRUE
      )
      
    }, error = function(e) {
      return(list(
        error = paste("Error in Welch ANOVA:", e$message),
        success = FALSE
      ))
    })
  })
  
  output$welch_anova_results <- renderPrint({
    req(welch_anova_results())
    
    results <- welch_anova_results()
    
    if (!results$success) {
      cat("WELCH ANOVA - ERROR\n")
      cat("===================\n\n")
      cat(results$error, "\n\n")
      cat("Troubleshooting suggestions:\n")
      cat("1. Check that your grouping variable has at least 2 categories\n")
      cat("2. Ensure each group has at least 2 observations\n")
      cat("3. Remove groups with zero variance (all identical values)\n")
      cat("4. Check for extreme outliers\n")
      cat("5. Try using Kruskal-Wallis test for non-normal data\n")
      return()
    }
    
    cat("WELCH ANOVA RESULTS\n")
    cat("===================\n\n")
    
    cat("Data Summary:\n")
    cat("- Outcome variable:", input$welch_outcome, "\n")
    cat("- Grouping variable:", input$welch_group, "\n")
    cat("- Number of groups:", results$n_groups, "\n")
    cat("- Total observations:", results$total_n, "\n")
    cat("\n")
    
    print(results$welch_test)
    
    cat("\n")
    
    # Add effect size (omega squared)
    if (!is.nan(results$welch_test$statistic)) {
      F_val <- results$welch_test$statistic
      df1 <- results$welch_test$parameter[1]
      df2 <- results$welch_test$parameter[2]
      
      # Calculate omega squared for Welch ANOVA (approximation)
      omega_sq <- (F_val - 1) / (F_val + (df2 + 1)/df1)
      eta_sq <- F_val * df1 / (F_val * df1 + df2)
      
      cat("Effect Sizes:\n")
      cat("- Omega squared (ω²):", round(omega_sq, 3), "\n")
      cat("- Eta squared (η²):", round(eta_sq, 3), "\n")
      cat("\n")
      
      # Interpretation
      cat("Effect Size Interpretation:\n")
      if (omega_sq >= 0.01 && omega_sq < 0.06) {
        cat("- Small effect\n")
      } else if (omega_sq >= 0.06 && omega_sq < 0.14) {
        cat("- Medium effect\n")
      } else if (omega_sq >= 0.14) {
        cat("- Large effect\n")
      } else {
        cat("- Negligible effect\n")
      }
    }
  })
  
  output$welch_anova_table <- renderDT({
    req(welch_anova_results())
    
    results <- welch_anova_results()
    
    if (!results$success) {
      return(datatable(
        data.frame(Error = results$error),
        options = list(dom = 't'),
        rownames = FALSE,
        caption = "Welch ANOVA - Error"
      ))
    }
    
    welch_df <- data.frame(
      Statistic = c("F value", "Numerator DF", "Denominator DF", "P-value", 
                    "Total Observations", "Number of Groups", "Method"),
      Value = c(
        round(results$welch_test$statistic, 4),
        round(results$welch_test$parameter[1], 2),
        round(results$welch_test$parameter[2], 2),
        round(results$welch_test$p.value, 4),
        results$total_n,
        results$n_groups,
        "Welch ANOVA (unequal variances)"
      )
    )
    
    datatable(
      welch_df,
      options = list(
        dom = 't',
        pageLength = 10
      ),
      rownames = FALSE,
      caption = "Welch ANOVA Results"
    ) %>%
      formatStyle(
        'Value',
        backgroundColor = styleInterval(
          c(0.001, 0.01, 0.05),
          c('#D4EDDA', '#D1ECF1', '#FFF3CD', '#FFFFFF')
        ),
        target = 'row',
        columns = 'Value'
      )
  })
  
  output$welch_posthoc_results <- renderPrint({
    req(welch_anova_results())
    
    results <- welch_anova_results()
    
    if (!results$success) {
      cat("Cannot perform post-hoc tests due to ANOVA error\n")
      return()
    }
    
    if (results$n_groups <= 2) {
      cat("Post-hoc tests are only performed when there are more than 2 groups.\n")
      return()
    }
    
    if (is.null(results$posthoc)) {
      cat("Post-hoc tests could not be calculated.\n")
      cat("Alternative: Try using pairwise Wilcoxon tests with Bonferroni correction.\n")
      return()
    }
    
    cat("POST-HOC COMPARISONS (Holm correction for unequal variances)\n")
    cat("============================================================\n\n")
    
    print(results$posthoc)
    
    # Add interpretation
    cat("\nInterpretation:\n")
    cat("- P-value < 0.05 indicates statistically significant difference between groups\n")
    cat("- Holm correction controls for multiple comparisons\n")
    cat("- This method is robust to unequal variances and non-normality\n")
  })
  
  # Additional output for group statistics
  output$welch_group_stats <- renderDT({
    req(welch_anova_results())
    
    results <- welch_anova_results()
    
    if (!results$success) {
      return(NULL)
    }
    
    datatable(
      results$group_stats,
      options = list(
        pageLength = 10,
        dom = 'Blfrtip',
        scrollX = TRUE,
        buttons = c('copy', 'csv', 'excel')
      ),
      rownames = FALSE,
      caption = "Group Descriptive Statistics",
      extensions = 'Buttons'
    ) %>%
      formatRound(columns = c('Mean', 'SD', 'SE', 'Variance'), digits = 3)
  })
  
  # Download handler for Welch ANOVA
  output$download_welch_anova <- downloadHandler(
    filename = function() {
      paste("welch_anova_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      req(welch_anova_results())
      results <- welch_anova_results()
      
      doc <- read_docx()
      
      if (!results$success) {
        doc <- doc %>%
          body_add_par("WELCH ANOVA ANALYSIS - ERROR", style = "heading 1") %>%
          body_add_par(results$error, style = "Normal")
      } else {
        doc <- doc %>%
          body_add_par("WELCH ANOVA ANALYSIS REPORT", style = "heading 1") %>%
          body_add_par(paste("Outcome Variable:", input$welch_outcome), style = "Normal") %>%
          body_add_par(paste("Grouping Variable:", input$welch_group), style = "Normal") %>%
          body_add_par("Method: Welch ANOVA (robust to unequal variances)", style = "Normal") %>%
          body_add_par(paste("Number of Groups:", results$n_groups), style = "Normal") %>%
          body_add_par(paste("Total Observations:", results$total_n), style = "Normal") %>%
          body_add_par("", style = "Normal")
        
        doc <- doc %>%
          body_add_par("GROUP DESCRIPTIVE STATISTICS", style = "heading 2")
        
        ft_stats <- flextable(results$group_stats) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_stats)
        
        doc <- doc %>%
          body_add_par("WELCH ANOVA RESULTS", style = "heading 2")
        
        welch_df <- data.frame(
          Statistic = c("F value", "Numerator DF", "Denominator DF", "P-value"),
          Value = c(
            round(results$welch_test$statistic, 4),
            round(results$welch_test$parameter[1], 2),
            round(results$welch_test$parameter[2], 2),
            round(results$welch_test$p.value, 4)
          )
        )
        
        ft_welch <- flextable(welch_df) %>%
          theme_zebra() %>%
          autofit()
        
        doc <- doc %>% body_add_flextable(ft_welch)
        
        # Calculate and add effect sizes
        if (!is.nan(results$welch_test$statistic)) {
          F_val <- results$welch_test$statistic
          df1 <- results$welch_test$parameter[1]
          df2 <- results$welch_test$parameter[2]
          
          omega_sq <- (F_val - 1) / (F_val + (df2 + 1)/df1)
          eta_sq <- F_val * df1 / (F_val * df1 + df2)
          
          effect_sizes <- data.frame(
            Effect_Size = c("Omega squared (ω²)", "Eta squared (η²)"),
            Value = c(round(omega_sq, 3), round(eta_sq, 3))
          )
          
          doc <- doc %>%
            body_add_par("EFFECT SIZES", style = "heading 2")
          
          ft_effects <- flextable(effect_sizes) %>%
            theme_zebra() %>%
            autofit()
          
          doc <- doc %>% body_add_flextable(ft_effects)
        }
        
        if (!is.null(results$posthoc)) {
          doc <- doc %>%
            body_add_par("POST-HOC COMPARISONS", style = "heading 2")
          
          posthoc_matrix <- as.data.frame(results$posthoc$p.value)
          posthoc_df <- cbind(Comparison = rownames(posthoc_matrix), posthoc_matrix)
          rownames(posthoc_df) <- NULL
          
          ft_posthoc <- flextable(posthoc_df) %>%
            theme_zebra() %>%
            autofit()
          
          doc <- doc %>% body_add_flextable(ft_posthoc)
        }
        
        doc <- doc %>%
          body_add_par("INTERPRETATION", style = "heading 2")
        
        p_value <- results$welch_test$p.value
        f_value <- results$welch_test$statistic
        
        interpretation <- if (p_value < 0.05) {
          paste("The Welch ANOVA revealed a statistically significant difference between groups ",
                "(F =", round(f_value, 3), ", p =", round(p_value, 4), "). ",
                "This robust test confirms significant differences while accounting for potential unequal variances.")
        } else {
          paste("The Welch ANOVA did not reveal a statistically significant difference between groups ",
                "(F =", round(f_value, 3), ", p =", round(p_value, 4), ").")
        }
        
        doc <- doc %>% body_add_par(interpretation, style = "Normal")
        
        doc <- doc %>%
          body_add_par("RECOMMENDATIONS", style = "heading 2") %>%
          body_add_par("• Use Welch ANOVA when group variances are unequal", style = "Normal") %>%
          body_add_par("• Consider Kruskal-Wallis test for non-normal data", style = "Normal") %>%
          body_add_par("• For post-hoc comparisons, use Games-Howell or pairwise t-tests with Holm correction", style = "Normal")
      }
      
      print(doc, target = file)
    }
  )
  # ANOVA Assumptions
  assumptions_results <- eventReactive(input$run_assumptions, {
    req(input$assumptions_outcome, input$assumptions_group, processed_data())
    df <- processed_data()
    
    # Fit ANOVA model
    model <- aov(as.formula(paste(input$assumptions_outcome, "~", input$assumptions_group)), data = df)
    
    results <- list()
    
    if(input$levene_test) {
      results$levene <- car::leveneTest(model)
    }
    
    if(input$normality_test) {
      results$shapiro <- shapiro.test(residuals(model))
      results$ks <- ks.test(residuals(model), "pnorm")
    }
    
    if(input$outlier_test) {
      residuals_std <- rstandard(model)
      results$outliers <- which(abs(residuals_std) > 2.5)
    }
    
    results
  })
  
  output$levene_results <- renderPrint({
    req(assumptions_results()$levene)
    
    cat("Levene's Test for Homogeneity of Variance\n")
    cat("=========================================\n\n")
    print(assumptions_results()$levene)
  })
  
  output$normality_results <- renderPrint({
    req(assumptions_results()$shapiro)
    
    cat("Normality Tests\n")
    cat("===============\n\n")
    
    cat("Shapiro-Wilk Test:\n")
    print(assumptions_results()$shapiro)
    
    if(!is.null(assumptions_results()$ks)) {
      cat("\nKolmogorov-Smirnov Test:\n")
      print(assumptions_results()$ks)
    }
  })
  
  output$outlier_results <- renderPrint({
    req(assumptions_results()$outliers)
    
    cat("Outlier Detection\n")
    cat("=================\n\n")
    
    if(length(assumptions_results()$outliers) > 0) {
      cat("Outliers detected at observations:", assumptions_results()$outliers, "\n")
    } else {
      cat("No significant outliers detected (standardized residuals < |2.5|)\n")
    }
  })
  
  output$assumptions_plots <- renderPlot({
    req(input$assumptions_outcome, input$assumptions_group, processed_data())
    
    df <- processed_data()
    model <- aov(as.formula(paste(input$assumptions_outcome, "~", input$assumptions_group)), data = df)
    
    par(mfrow = c(2, 2))
    plot(model)
  })
  
  # One-Sample T-Test
  ttest_onesample_results <- eventReactive(input$run_ttest_onesample, {
    req(input$ttest_onesample_var, processed_data())
    df <- processed_data()
    
    var_data <- df[[input$ttest_onesample_var]]
    var_data <- var_data[!is.na(var_data)]
    
    # Perform t-test
    t_test <- t.test(var_data, 
                     mu = input$ttest_onesample_mu,
                     alternative = input$ttest_onesample_alternative,
                     conf.level = input$ttest_onesample_conf)
    
    list(
      ttest = t_test,
      data = var_data,
      descriptives = data.frame(
        N = length(var_data),
        Mean = mean(var_data),
        SD = sd(var_data),
        SE = sd(var_data) / sqrt(length(var_data))
      )
    )
  })
  
  output$ttest_onesample_results <- renderPrint({
    req(ttest_onesample_results())
    
    cat("One-Sample T-Test Results\n")
    cat("=========================\n\n")
    
    print(ttest_onesample_results()$ttest)
  })
  
  output$ttest_onesample_table <- renderDT({
    req(ttest_onesample_results())
    
    ttest_df <- data.frame(
      Statistic = c("t", "DF", "P-value", "Confidence Level", "Alternative"),
      Value = c(
        round(ttest_onesample_results()$ttest$statistic, 3),
        round(ttest_onesample_results()$ttest$parameter, 2),
        round(ttest_onesample_results()$ttest$p.value, 4),
        paste0(input$ttest_onesample_conf * 100, "%"),
        ttest_onesample_results()$ttest$alternative
      )
    )
    
    datatable(ttest_df, options = list(dom = 't'))
  })
  
  output$ttest_onesample_descriptives <- renderPrint({
    req(ttest_onesample_results())
    
    cat("Descriptive Statistics\n")
    cat("=====================\n\n")
    
    print(ttest_onesample_results()$descriptives)
  })
  
  output$ttest_onesample_plot <- renderPlot({
    req(ttest_onesample_results())
    
    data <- ttest_onesample_results()$data
    mu <- input$ttest_onesample_mu
    
    ggplot(data.frame(value = data), aes(x = value)) +
      geom_histogram(aes(y = ..density..), bins = 30, fill = "#0C5EA8", alpha = 0.7) +
      geom_density(color = "#CD2026", size = 1.5) +
      geom_vline(xintercept = mu, color = "#5CB85C", size = 1.5, linetype = "dashed") +
      labs(title = "Distribution with Test Value",
           x = input$ttest_onesample_var,
           y = "Density") +
      theme_minimal() +
      theme(
        plot.title = element_text(size = 16, face = "bold", hjust = 0.5),
        axis.title = element_text(size = 14, face = "bold"),
        axis.text = element_text(size = 12)
      )
  })
  
  # UI for independent t-test group levels
  output$ttest_independent_levels_ui <- renderUI({
    req(input$ttest_independent_group, processed_data())
    df <- processed_data()
    
    if (!input$ttest_independent_group %in% names(df)) return(NULL)
    
    levels <- unique(na.omit(df[[input$ttest_independent_group]]))
    
    tagList(
      selectInput("ttest_group1", "Group 1:", choices = levels, selected = levels[1]),
      selectInput("ttest_group2", "Group 2:", choices = levels, selected = levels[2])
    )
  })
  
  # Independent T-Test
  ttest_independent_results <- eventReactive(input$run_ttest_independent, {
    req(input$ttest_independent_var, input$ttest_independent_group, 
        input$ttest_group1, input$ttest_group2, processed_data())
    
    df <- processed_data()
    
    # Filter data for the two selected groups
    group1_data <- df[[input$ttest_independent_var]][df[[input$ttest_independent_group]] == input$ttest_group1]
    group2_data <- df[[input$ttest_independent_var]][df[[input$ttest_independent_group]] == input$ttest_group2]
    
    group1_data <- group1_data[!is.na(group1_data)]
    group2_data <- group2_data[!is.na(group2_data)]
    
    # Perform t-test
    t_test <- t.test(group1_data, group2_data,
                     var.equal = input$ttest_var_equal,
                     alternative = input$ttest_alternative,
                     conf.level = input$ttest_conf)
    
    # Variance test
    var_test <- var.test(group1_data, group2_data)
    
    list(
      ttest = t_test,
      var_test = var_test,
      descriptives = data.frame(
        Group = c(input$ttest_group1, input$ttest_group2),
        N = c(length(group1_data), length(group2_data)),
        Mean = c(mean(group1_data), mean(group2_data)),
        SD = c(sd(group1_data), sd(group2_data)),
        SE = c(sd(group1_data)/sqrt(length(group1_data)), 
               sd(group2_data)/sqrt(length(group2_data)))
      )
    )
  })
  
  output$ttest_independent_results <- renderPrint({
    req(ttest_independent_results())
    
    cat("Independent T-Test Results\n")
    cat("==========================\n\n")
    
    print(ttest_independent_results()$ttest)
  })
  
  output$ttest_independent_table <- renderDT({
    req(ttest_independent_results())
    
    ttest_df <- data.frame(
      Statistic = c("t", "DF", "P-value", "Confidence Level", "Alternative", "Equal Variances"),
      Value = c(
        round(ttest_independent_results()$ttest$statistic, 3),
        round(ttest_independent_results()$ttest$parameter, 2),
        round(ttest_independent_results()$ttest$p.value, 4),
        paste0(input$ttest_conf * 100, "%"),
        ttest_independent_results()$ttest$alternative,
        ifelse(input$ttest_var_equal, "Yes", "No")
      )
    )
    
    datatable(ttest_df, options = list(dom = 't'))
  })
  
  output$ttest_independent_descriptives <- renderPrint({
    req(ttest_independent_results())
    
    cat("Group Descriptive Statistics\n")
    cat("===========================\n\n")
    
    print(ttest_independent_results()$descriptives)
  })
  
  output$ttest_variance_check <- renderPrint({
    req(ttest_independent_results())
    
    cat("Variance Equality Test (F-test)\n")
    cat("===============================\n\n")
    
    print(ttest_independent_results()$var_test)
  })
  
  output$ttest_independent_plot <- renderPlot({
    req(ttest_independent_results())
    
    desc <- ttest_independent_results()$descriptives
    
    ggplot(desc, aes(x = Group, y = Mean, fill = Group)) +
      geom_bar(stat = "identity", alpha = 0.7) +
      geom_errorbar(aes(ymin = Mean - 1.96*SE, ymax = Mean + 1.96*SE), 
                    width = 0.2) +
      scale_fill_brewer(palette = "Set1") +
      labs(title = "Group Means with 95% Confidence Intervals",
           x = input$ttest_independent_group,
           y = input$ttest_independent_var) +
      theme_minimal() +
      theme(
        plot.title = element_text(size = 16, face = "bold", hjust = 0.5),
        axis.title = element_text(size = 14, face = "bold"),
        axis.text = element_text(size = 12),
        legend.position = "none"
      )
  })
  
  # Paired T-Test
  ttest_paired_results <- eventReactive(input$run_ttest_paired, {
    req(input$ttest_paired_var1, input$ttest_paired_var2, processed_data())
    df <- processed_data()
    
    var1_data <- df[[input$ttest_paired_var1]]
    var2_data <- df[[input$ttest_paired_var2]]
    
    # Remove pairs with missing values
    complete_cases <- !is.na(var1_data) & !is.na(var2_data)
    var1_clean <- var1_data[complete_cases]
    var2_clean <- var2_data[complete_cases]
    
    # Perform paired t-test
    t_test <- t.test(var1_clean, var2_clean,
                     paired = TRUE,
                     alternative = input$ttest_paired_alternative,
                     conf.level = input$ttest_paired_conf)
    
    list(
      ttest = t_test,
      var1_data = var1_clean,
      var2_data = var2_clean,
      differences = var1_clean - var2_clean,
      descriptives = data.frame(
        Variable = c(input$ttest_paired_var1, input$ttest_paired_var2),
        N = c(length(var1_clean), length(var2_clean)),
        Mean = c(mean(var1_clean), mean(var2_clean)),
        SD = c(sd(var1_clean), sd(var2_clean)),
        SE = c(sd(var1_clean)/sqrt(length(var1_clean)), 
               sd(var2_clean)/sqrt(length(var2_clean)))
      )
    )
  })
  
  output$ttest_paired_results <- renderPrint({
    req(ttest_paired_results())
    
    cat("Paired T-Test Results\n")
    cat("====================\n\n")
    
    print(ttest_paired_results()$ttest)
  })
  
  output$ttest_paired_table <- renderDT({
    req(ttest_paired_results())
    
    ttest_df <- data.frame(
      Statistic = c("t", "DF", "P-value", "Confidence Level", "Alternative"),
      Value = c(
        round(ttest_paired_results()$ttest$statistic, 3),
        round(ttest_paired_results()$ttest$parameter, 2),
        round(ttest_paired_results()$ttest$p.value, 4),
        paste0(input$ttest_paired_conf * 100, "%"),
        ttest_paired_results()$ttest$alternative
      )
    )
    
    datatable(ttest_df, options = list(dom = 't'))
  })
  
  output$ttest_paired_descriptives <- renderPrint({
    req(ttest_paired_results())
    
    cat("Descriptive Statistics for Paired Variables\n")
    cat("==========================================\n\n")
    
    print(ttest_paired_results()$descriptives)
  })
  
  output$ttest_paired_differences <- renderPrint({
    req(ttest_paired_results())
    
    diff_data <- ttest_paired_results()$differences
    
    cat("Difference Analysis (Var1 - Var2)\n")
    cat("=================================\n\n")
    
    cat("Mean difference:", round(mean(diff_data), 3), "\n")
    cat("SD of differences:", round(sd(diff_data), 3), "\n")
    cat("95% CI for mean difference: [", 
        round(mean(diff_data) - 1.96*sd(diff_data)/sqrt(length(diff_data)), 3),
        ", ",
        round(mean(diff_data) + 1.96*sd(diff_data)/sqrt(length(diff_data)), 3),
        "]\n")
  })
  
  output$ttest_paired_plot <- renderPlot({
    req(ttest_paired_results())
    
    desc <- ttest_paired_results()$descriptives
    diff_data <- ttest_paired_results()$differences
    
    plot_data <- data.frame(
      Type = rep(c("Individual", "Difference"), each = length(diff_data)),
      Value = c(ttest_paired_results()$var1_data, 
                ttest_paired_results()$var2_data,
                diff_data)
    )
    
    ggplot(plot_data, aes(x = Type, y = Value, fill = Type)) +
      geom_boxplot(alpha = 0.7) +
      scale_fill_brewer(palette = "Set2") +
      labs(title = "Boxplot of Paired Variables and Their Differences",
           x = "",
           y = "Value") +
      theme_minimal() +
      theme(
        plot.title = element_text(size = 16, face = "bold", hjust = 0.5),
        axis.title = element_text(size = 14, face = "bold"),
        axis.text = element_text(size = 12),
        legend.position = "none"
      )
  })
  
  # Mann-Whitney U Test
  mannwhitney_results <- eventReactive(input$run_mannwhitney, {
    req(input$mannwhitney_var, input$mannwhitney_group, processed_data())
    df <- processed_data()
    
    var_data <- df[[input$mannwhitney_var]]
    group_data <- df[[input$mannwhitney_group]]
    
    # Remove missing values
    complete_cases <- !is.na(var_data) & !is.na(group_data)
    var_clean <- var_data[complete_cases]
    group_clean <- group_data[complete_cases]
    
    # Perform Mann-Whitney U test
    test <- wilcox.test(var_clean ~ group_clean)
    
    list(test = test)
  })
  
  output$mannwhitney_results <- renderPrint({
    req(mannwhitney_results())
    
    cat("Mann-Whitney U Test Results\n")
    cat("===========================\n\n")
    
    print(mannwhitney_results()$test)
  })
  
  # Wilcoxon Signed-Rank Test
  wilcoxon_results <- eventReactive(input$run_wilcoxon, {
    req(input$wilcoxon_var1, input$wilcoxon_var2, processed_data())
    df <- processed_data()
    
    var1_data <- df[[input$wilcoxon_var1]]
    var2_data <- df[[input$wilcoxon_var2]]
    
    # Remove missing values
    complete_cases <- !is.na(var1_data) & !is.na(var2_data)
    var1_clean <- var1_data[complete_cases]
    var2_clean <- var2_data[complete_cases]
    
    # Perform Wilcoxon signed-rank test
    test <- wilcox.test(var1_clean, var2_clean, paired = TRUE)
    
    list(test = test)
  })
  
  output$wilcoxon_results <- renderPrint({
    req(wilcoxon_results())
    
    cat("Wilcoxon Signed-Rank Test Results\n")
    cat("=================================\n\n")
    
    print(wilcoxon_results()$test)
  })
  
  # Navigation button handlers
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
  
  # Helper function to format p-values
  format_p_value <- function(p) {
    if (is.na(p)) return("NA")
    if (p < 0.001) return("<0.001")
    return(sprintf("%.3f", p))
  }
  
  # Session info
  output$session_info <- renderPrint({
    sessionInfo()
  })
}

# Run the application
shinyApp(ui = ui, server = server)
