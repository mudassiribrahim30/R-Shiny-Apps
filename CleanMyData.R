library(shiny)
library(shinythemes)
library(DT)
library(tidyverse)
library(haven)
library(readxl)
library(shinyjs)
library(shinyWidgets)
library(bslib)
library(plotly)
library(VIM)
library(shinycssloaders)
library(DescTools)
library(zoo)
library(class)
library(fastDummies)

# Developer information
developer_info <- "Developed by Mudasir Mohammed Ibrahim. For any problems or suggestions, contact mudassiribrahim30@gmail.com"

ui <- fluidPage(
  theme = bslib::bs_theme(bootswatch = "minty"),
  useShinyjs(),
  tags$head(
    tags$style(HTML("
      .data-cleaning-section {
        background-color: #f8f9fa;
        border-radius: 8px;
        padding: 15px;
        margin-bottom: 20px;
      }
      .well {
        background-color: #ffffff;
      }
      .missing-cell {
        background-color: #ffcccc !important;
      }
      .developer-footer {
        text-align: center;
        margin-top: 20px;
        font-size: 12px;
        color: #666;
      }
      .help-section {
        background-color: #e8f4fc;
        border-radius: 8px;
        padding: 15px;
        margin-bottom: 20px;
      }
      
      /* New styles for scrollable sidebar */
      .sidebar {
        height: 100vh;
        overflow-y: auto;
        position: sticky;
        top: 0;
      }
      
      .main-panel {
        height: 100vh;
        overflow-y: auto;
      }
      
      .data-cleaning-section {
        margin-bottom: 15px;
        padding-bottom: 15px;
      }
      
      .range-recode-box {
        border: 1px solid #ddd;
        border-radius: 5px;
        padding: 10px;
        margin-bottom: 10px;
        background-color: #f9f9f9;
      }
    "))
  ),
  
  titlePanel(
    div(
      "CleanMyData",
      style = "color: #2c3e50; font-weight: bold;"
    ),
    windowTitle = "Data Cleaning App"
  ),
  
  sidebarLayout(
    sidebarPanel(
      width = 3,
      class = "sidebar",
      # File upload section
      div(
        class = "data-cleaning-section",
        h4("Data Import", style = "color: #2c3e50;"),
        fileInput(
          "file", 
          NULL,
          buttonLabel = "Browse...",
          placeholder = "CSV, SPSS, Stata, Excel",
          accept = c(".csv", ".sav", ".dta", ".xlsx", ".xls")
        ),
        radioGroupButtons(
          "header_option", 
          "Header Option",
          choices = c("First row" = "first_row", "Specific row" = "specific_row", "No header" = "no_header"),
          selected = "first_row",
          status = "primary"
        ),
        conditionalPanel(
          condition = "input.header_option == 'specific_row'",
          numericInput("header_row", "Row number for column names", value = 1, min = 1),
          actionBttn(
            "apply_header",
            "Apply Header",
            style = "material-flat",
            color = "primary",
            size = "sm"
          )
        ),
        uiOutput("file_options_ui")
      ),
      
      # Data operations section
      div(
        class = "data-cleaning-section",
        h4("Data Operations", style = "color: #2c3e50;"),
        actionBttn(
          "transpose", 
          "Transpose Data", 
          icon = icon("exchange-alt"),
          style = "gradient",
          color = "primary",
          block = TRUE
        ),
        actionBttn(
          "reset", 
          "Reset Data", 
          icon = icon("undo"),
          style = "simple",
          color = "warning",
          block = TRUE
        )
      ),
      
      # Data cleaning section
      div(
        class = "data-cleaning-section",
        h4("Data Cleaning", style = "color: #2c3e50;"),
        pickerInput(
          "clean_op", 
          "Operation",
          choices = c("Data Editing", "Missing Data", "Outliers", "Imputation"),
          selected = "Data Editing",
          multiple = FALSE,
          options = list(style = "btn-danger")
        ),
        uiOutput("clean_op_ui")
      ),
      
      # Column operations section
      div(
        class = "data-cleaning-section",
        h4("Column Operations", style = "color: #2c3e50;"),
        pickerInput(
          "col_op", 
          "Operation",
          choices = c("Rename", "Recode", "Create", "Delete", "Type Conversion", "Row/Column Sums", "Dummy Code", "Column Arrangement"),
          multiple = FALSE,
          options = list(style = "btn-primary")
        ),
        uiOutput("col_op_ui")
      ),
      
      # Row operations section
      div(
        class = "data-cleaning-section",
        h4("Row Operations", style = "color: #2c3e50;"),
        pickerInput(
          "row_op", 
          "Operation",
          choices = c("Filter", "Sort", "Remove Duplicates", "Sample"),
          multiple = FALSE,
          options = list(style = "btn-info")
        ),
        uiOutput("row_op_ui")
      ),
      
      # Descriptive statistics section
      div(
        class = "data-cleaning-section",
        h4("Descriptive Statistics", style = "color: #2c3e50;"),
        pickerInput(
          "desc_var", 
          "Select Variable",
          choices = NULL,
          multiple = FALSE
        ),
        actionBttn(
          "run_desc", 
          "Run Analysis", 
          icon = icon("calculator"),
          style = "material-flat",
          color = "primary",
          size = "sm"
        )
      ),
      
      # Download section
      div(
        style = "margin-top: 30px;",
        pickerInput(
          "download_format",
          "Download Format",
          choices = c("CSV", "Excel", "SPSS", "Stata"),
          selected = "CSV",
          multiple = FALSE
        ),
        pickerInput(
          "download_cols",
          "Columns to Export (Leave empty for all)",
          choices = NULL,
          multiple = TRUE,
          options = list(
            `actions-box` = TRUE,
            `selected-text-format` = "count > 2"
          )
        ),
        downloadBttn(
          "downloadData", 
          "Download Data",
          style = "bordered",
          color = "success",
          size = "md",
          block = TRUE
        )
      ),
      
      # Developer info
      div(
        class = "developer-footer",
        developer_info
      )
    ),
    
    mainPanel(
      width = 9,
      class = "main-panel",
      tabsetPanel(
        type = "tabs",
        tabPanel(
          "Data View",
          icon = icon("table"),
          div(
            style = "margin-top: 10px;",
            DTOutput("data_table") %>% 
              withSpinner(color = "#18bc9c")
          )
        ),
        
        tabPanel(
          "Summary",
          icon = icon("chart-bar"),
          verbatimTextOutput("summary") %>% 
            withSpinner(color = "#18bc9c")
        ),
        tabPanel(
          "Missing Data",
          icon = icon("search"),
          plotlyOutput("missing_plot") %>% 
            withSpinner(color = "#18bc9c"),
          verbatimTextOutput("missing_summary") %>% 
            withSpinner(color = "#18bc9c")
        ),
        tabPanel(
          "Structure",
          icon = icon("project-diagram"),
          verbatimTextOutput("structure") %>% 
            withSpinner(color = "#18bc9c")
        ),
        tabPanel(
          "Plots",
          icon = icon("image"),
          selectInput("plot_type", "Plot Type",
                      choices = c("Histogram", "Bar Plot", "Box Plot", "Scatter Plot"),
                      selected = "Histogram"),
          uiOutput("plot_options_ui"),
          plotOutput("data_plot") %>% 
            withSpinner(color = "#18bc9c")
        ),
        tabPanel(
          "Descriptive Stats",
          icon = icon("calculator"),
          verbatimTextOutput("descriptive_stats") %>% 
            withSpinner(color = "#18bc9c")
        ),
        tabPanel(
          "Help & Instructions",
          icon = icon("question-circle"),
          div(
            class = "help-section",
            h3("How to Use This App", style = "color: #2c3e50;"),
            h4("1. Data Import", style = "color: #3498db;"),
            p("Upload your data file (CSV, Excel, SPSS, or Stata format) using the 'Browse' button."),
            p("For CSV files, you can specify the separator and decimal options."),
            p("Choose header option:"),
            p("- First row: Use first row as column names"),
            p("- Specific row: Select which row to use as column names (click 'Apply Header' to apply)"),
            p("- No header: Don't use any row as column names"),
            
            h4("2. Data Operations", style = "color: #3498db;"),
            p("- Transpose Data: Switch rows and columns"),
            p("- Reset Data: Revert to the original uploaded data"),
            
            h4("3. Data Cleaning", style = "color: #3498db;"),
            p("- Missing Data: Highlight and analyze missing values"),
            p("- Outliers: Detect and remove outliers based on standard deviations"),
            p("- Data Editing: Double-click cells to edit values directly. After making changes, click 'Save Edits' to apply them."),
            p("- Imputation: Fill missing values using various methods (Mean, Median, Mode, Zero, Linear Interpolation, KNN)"),
            
            h4("4. Column Operations", style = "color: #3498db;"),
            p("- Rename: Change column names"),
            p("- Recode: Modify values in a column (including numerical to categorical with ranges)"),
            p("- Create: Add new columns with specified types"),
            p("- Delete: Remove columns"),
            p("- Type Conversion: Change column data types"),
            p("- Row/Column Sums: Calculate sums across rows or columns for selected variables"),
            p("- Dummy Code: Convert categorical variables to dummy variables (0/1 coding)"),
            p("- Column Arrangement: Reorder columns as desired"),
            
            h4("5. Row Operations", style = "color: #3498db;"),
            p("- Filter: Subset data based on conditions"),
            p("- Sort: Order data by selected columns"),
            p("- Remove Duplicates: Eliminate duplicate rows (either entire rows or based on specific columns)"),
            p("- Sample: Take a random sample of rows"),
            
            h4("6. Descriptive Statistics", style = "color: #3498db;"),
            p("Select a variable and click 'Run Analysis' to get detailed statistics"),
            p("For numeric variables: mean, SD, median, mode, min, max, skewness, kurtosis"),
            p("For categorical variables: frequency, percentage, median, mode"),
            
            h4("7. Data Export", style = "color: #3498db;"),
            p("Download your processed data in CSV, Excel, SPSS, or Stata format"),
            p("You can select specific columns to export using the 'Columns to Export' picker"),
            
            h4("Recoding Variables", style = "color: #3498db;"),
            p("To recode variables:"),
            p("1. Select 'Recode' in Column Operations"),
            p("2. Choose the column to recode"),
            p("3. For numerical variables, you can specify ranges with min/max values and labels"),
            p("4. For categorical variables, specify exact value mappings"),
            p("5. Choose whether to create a new variable or replace existing"),
            p("6. Click 'Recode' to apply the changes"),
            
            h4("Dummy Coding", style = "color: #3498db;"),
            p("To create dummy variables:"),
            p("1. Select 'Dummy Code' in Column Operations"),
            p("2. Choose the categorical column to convert"),
            p("3. Select whether to keep the original column"),
            p("4. Click 'Create Dummies' to generate the new variables"),
            
            h4("Column Arrangement", style = "color: #3498db;"),
            p("To rearrange columns:"),
            p("1. Select 'Column Arrangement' in Column Operations"),
            p("2. Drag and drop columns to reorder them as desired"),
            p("3. Click 'Apply Order' to save the new column arrangement"),
            
            h4("Calculating Row/Column Sums", style = "color: #3498db;"),
            p("To calculate sums across rows or columns:"),
            p("1. Select 'Row/Column Sums' in Column Operations"),
            p("2. Choose the variables to include in the calculation"),
            p("3. Select whether to calculate row sums or column sums"),
            p("4. Enter a name for the new variable that will contain the sums"),
            p("5. Click 'Calculate Sums' to apply the operation"),
            
            h4("Removing Duplicates", style = "color: #3498db;"),
            p("You can remove duplicates by:"),
            p("- Entire rows (all columns must match)"),
            p("- Specific columns (only selected columns must match)"),
            
            h4("Data Editing", style = "color: #3498db;"),
            p("To edit data directly:"),
            p("1. Double-click on any cell to edit its value"),
            p("2. Make your changes"),
            p("3. Click 'Save Edits' to apply the changes permanently"),
            p("Note: After applying other operations (like filtering or sorting), you can still edit data by selecting 'Data Editing' in the Data Cleaning section"),
            
            h4("Troubleshooting", style = "color: #3498db;"),
            p("- If operations don't work, check that you've selected the correct columns"),
            p("- For numerical operations, ensure columns are of numeric type"),
            p("- If the app freezes, try resetting the data or refreshing the page")
          )
        )
      )
    )
  )
)

server <- function(input, output, session) {
  
  # Reactive value to store data
  rv <- reactiveValues(data = NULL, original = NULL, last_saved = NULL, edited_cells = NULL)
  
  # Update variable selection for descriptive stats and download columns
  observe({
    req(rv$data)
    updatePickerInput(session, "desc_var", choices = names(rv$data))
    updatePickerInput(session, "download_cols", choices = names(rv$data))
  })
  
  # Dynamic UI for file import options
  output$file_options_ui <- renderUI({
    req(input$file)
    ext <- tools::file_ext(input$file$name)
    
    if (ext == "csv") {
      tagList(
        radioGroupButtons(
          "sep", 
          "Separator",
          choices = c(Comma = ",", Semicolon = ";", Tab = "\t"),
          selected = ",",
          status = "primary"
        ),
        radioGroupButtons(
          "decimal", 
          "Decimal",
          choices = c(Point = ".", Comma = ","),
          selected = ".",
          status = "primary"
        )
      )
    } else {
      NULL
    }
  })
  
  # Dynamic UI for plot options
  output$plot_options_ui <- renderUI({
    req(rv$data)
    
    if (input$plot_type == "Histogram") {
      selectInput("hist_col", "Select Numeric Column",
                  choices = names(select(rv$data, where(is.numeric))))
      
    } else if (input$plot_type == "Bar Plot") {
      selectInput("bar_col", "Select Categorical Column",
                  choices = names(select(rv$data, where(~!is.numeric(.)))))
      
    } else if (input$plot_type == "Box Plot") {
      tagList(
        selectInput("box_col", "Select Numeric Column",
                    choices = names(select(rv$data, where(is.numeric)))),
        selectInput("box_group", "Group By (Optional)",
                    choices = c("None", names(select(rv$data, where(~!is.numeric(.))))))
      )
      
    } else if (input$plot_type == "Scatter Plot") {
      tagList(
        selectInput("scatter_x", "X Axis",
                    choices = names(select(rv$data, where(is.numeric)))),
        selectInput("scatter_y", "Y Axis",
                    choices = names(select(rv$data, where(is.numeric)))),
        selectInput("scatter_color", "Color By (Optional)",
                    choices = c("None", names(select(rv$data, where(~!is.numeric(.))))))
      )
    }
  })
  
  # Dynamic UI for data cleaning operations
  output$clean_op_ui <- renderUI({
    req(input$clean_op, rv$data)
    
    switch(
      input$clean_op,
      "Missing Data" = tagList(
        h5("Missing Data Summary"),
        verbatimTextOutput("missing_stats"),
        pickerInput(
          "missing_cols", 
          "Columns to Analyze",
          choices = names(rv$data),
          multiple = TRUE,
          options = list(
            `actions-box` = TRUE,
            `selected-text-format` = "count > 2"
          )
        ),
        actionBttn(
          "show_missing",
          "Highlight Missing",
          style = "material-flat",
          color = "warning",
          size = "sm"
        )
      ),
      "Outliers" = tagList(
        selectInput(
          "outlier_col", 
          "Select Numeric Column",
          choices = names(select(rv$data, where(is.numeric))),
          width = "100%"
        ),
        numericInput(
          "outlier_sd", 
          "Standard Deviation Threshold",
          value = 3,
          min = 1,
          step = 0.5
        ),
        actionBttn(
          "detect_outliers",
          "Detect Outliers",
          style = "material-flat",
          color = "warning",
          size = "sm"
        ),
        actionBttn(
          "remove_outliers",
          "Remove Outliers",
          style = "material-flat",
          color = "danger",
          size = "sm"
        )
      ),
      "Data Editing" = tagList(
        h5("Double-click cells to edit values"),
        p("After making changes, click 'Save Edits' to apply them."),
        actionBttn(
          "save_edits",
          "Save Edits",
          style = "material-flat",
          color = "success",
          size = "sm"
        )
      ),
      "Imputation" = tagList(
        pickerInput(
          "impute_cols", 
          "Columns to Impute",
          choices = names(select(rv$data, where(is.numeric))),
          multiple = TRUE,
          options = list(
            `actions-box` = TRUE,
            `selected-text-format` = "count > 2"
          )
        ),
        selectInput(
          "impute_method", 
          "Imputation Method",
          choices = c("Mean", "Median", "Mode", "Zero", "Linear Interpolation", "KNN"),
          width = "100%"
        ),
        conditionalPanel(
          condition = "input.impute_method == 'KNN'",
          numericInput(
            "k_value", 
            "Number of Neighbors (k)",
            value = 5,
            min = 1,
            step = 1
          )
        ),
        actionBttn(
          "apply_impute",
          "Apply Imputation",
          style = "material-flat",
          color = "primary",
          size = "sm"
        )
      )
    )
  })
  
  # Dynamic UI for column operations
  output$col_op_ui <- renderUI({
    req(input$col_op, rv$data)
    
    switch(
      input$col_op,
      "Rename" = tagList(
        selectInput(
          "rename_col", 
          "Select Column", 
          choices = names(rv$data),
          width = "100%"
        ),
        textInput(
          "new_name", 
          "New Name",
          placeholder = "Enter new column name"
        ),
        actionBttn(
          "apply_rename",
          "Rename",
          style = "material-flat",
          color = "primary",
          size = "sm"
        )
      ),
      "Recode" = tagList(
        selectInput(
          "recode_col", 
          "Select Column", 
          choices = names(rv$data),
          width = "100%"
        ),
        radioGroupButtons(
          "recode_type",
          "Recode Type",
          choices = c("Create New Variable" = "new", "Replace Existing" = "replace"),
          selected = "new",
          status = "primary"
        ),
        conditionalPanel(
          condition = "input.recode_type == 'new'",
          textInput(
            "new_recode_name",
            "New Variable Name",
            placeholder = "Enter name for new variable"
          )
        ),
        # Conditional UI for numerical vs categorical recoding
        uiOutput("recode_ui"),
        actionBttn(
          "apply_recode",
          "Recode",
          style = "material-flat",
          color = "primary",
          size = "sm"
        )
      ),
      "Create" = tagList(
        textInput(
          "new_col_name", 
          "New Column Name",
          placeholder = "Enter column name"
        ),
        selectInput(
          "new_col_type", 
          "Data Type",
          choices = c("Numeric", "Character", "Logical", "Date"),
          width = "100%"
        ),
        conditionalPanel(
          condition = "input.new_col_type == 'Numeric'",
          numericInput(
            "num_val", 
            "Initial Value",
            value = 0
          )
        ),
        conditionalPanel(
          condition = "input.new_col_type == 'Character'",
          textInput(
            "char_val", 
            "Initial Value",
            value = ""
          )
        ),
        actionBttn(
          "apply_create",
          "Create",
          style = "material-flat",
          color = "primary",
          size = "sm"
        )
      ),
      "Delete" = tagList(
        pickerInput(
          "delete_cols", 
          "Select Columns",
          choices = names(rv$data),
          multiple = TRUE,
          options = list(
            `actions-box` = TRUE,
            `selected-text-format` = "count > 2"
          )
        ),
        actionBttn(
          "apply_delete",
          "Delete Selected",
          style = "material-flat",
          color = "danger",
          size = "sm"
        )
      ),
      "Type Conversion" = tagList(
        selectInput(
          "convert_col", 
          "Select Column", 
          choices = names(rv$data),
          width = "100%"
        ),
        selectInput(
          "new_type", 
          "Convert To",
          choices = c("Numeric", "Character", "Factor", "Date"),
          width = "100%"
        ),
        actionBttn(
          "apply_convert",
          "Convert",
          style = "material-flat",
          color = "primary",
          size = "sm"
        )
      ),
      "Row/Column Sums" = tagList(
        pickerInput(
          "sum_vars", 
          "Select Variables to Sum",
          choices = names(select(rv$data, where(is.numeric))),
          multiple = TRUE,
          options = list(
            `actions-box` = TRUE,
            `selected-text-format` = "count > 2"
          )
        ),
        radioGroupButtons(
          "sum_type",
          "Sum Type",
          choices = c("Row Sum" = "row", "Column Sum" = "col"),
          selected = "row",
          status = "primary"
        ),
        textInput(
          "sum_name", 
          "New Variable Name",
          placeholder = "Enter name for sum variable"
        ),
        actionBttn(
          "apply_sum",
          "Calculate Sums",
          style = "material-flat",
          color = "primary",
          size = "sm"
        )
      ),
      "Dummy Code" = tagList(
        selectInput(
          "dummy_col", 
          "Select Categorical Column", 
          choices = names(select(rv$data, where(~!is.numeric(.)))),
          width = "100%"
        ),
        checkboxInput(
          "keep_dummy_col",
          "Keep Original Column",
          value = FALSE
        ),
        actionBttn(
          "apply_dummy",
          "Create Dummies",
          style = "material-flat",
          color = "primary",
          size = "sm"
        )
      ),
      "Column Arrangement" = tagList(
        pickerInput(
          "col_order", 
          "Drag to Reorder Columns",
          choices = names(rv$data),
          multiple = TRUE,
          options = list(
            `actions-box` = TRUE,
            `selected-text-format` = "count > 2",
            `live-search` = TRUE,
            `multiple-separator` = " | "
          ),
          selected = names(rv$data)
        ),
        actionBttn(
          "apply_col_order",
          "Apply Order",
          style = "material-flat",
          color = "primary",
          size = "sm"
        )
      )
    )
  })
  
  # Dynamic UI for recoding based on column type
  output$recode_ui <- renderUI({
    req(input$recode_col, rv$data)
    
    if (is.numeric(rv$data[[input$recode_col]])) {
      # Numerical variable - show range-based recoding
      tagList(
        h5("Range-based Recoding"),
        div(
          class = "range-recode-box",
          fluidRow(
            column(4, numericInput("recode_min", "Min Value", value = NA)),
            column(4, selectInput("recode_op", "Operator", choices = c("<", "<=", ">", ">=", "==", "!=", "between"))),
            column(4, numericInput("recode_max", "Max Value", value = NA))
          ),
          textInput("recode_label", "New Label", placeholder = "Enter label for this range"),
          actionBttn(
            "add_range",
            "Add Range",
            style = "material-flat",
            color = "primary",
            size = "xs"
          )
        ),
        textAreaInput(
          "recode_map", 
          "Recoding Rules",
          placeholder = "Ranges will appear here as: min-max=label\nExample: 1-5=Low,6-10=Medium,11-15=High",
          rows = 4
        )
      )
    } else {
      # Categorical variable - show exact value recoding
      textAreaInput(
        "recode_map", 
        "Recoding Rules",
        placeholder = "old=new (for exact values)\nExample: Male=1,Female=2",
        rows = 4
      )
    }
  })
  
  # Add range to recoding rules
  observeEvent(input$add_range, {
    req(input$recode_col, input$recode_op, input$recode_label, rv$data)
    
    if (input$recode_op == "between") {
      if (is.na(input$recode_min) || is.na(input$recode_max)) {
        showNotification("Please specify both min and max values for 'between' operator", type = "error")
        return()
      }
      new_rule <- paste0(input$recode_min, "-", input$recode_max, "=", input$recode_label)
    } else {
      if (is.na(input$recode_min)) {
        showNotification("Please specify a value", type = "error")
        return()
      }
      new_rule <- paste0(input$recode_op, input$recode_min, "=", input$recode_label)
    }
    
    current_rules <- isolate(input$recode_map)
    if (current_rules == "") {
      updateTextAreaInput(session, "recode_map", value = new_rule)
    } else {
      updateTextAreaInput(session, "recode_map", value = paste0(current_rules, ",", new_rule))
    }
  })
  
  # Calculate row or column sums
  observeEvent(input$apply_sum, {
    req(rv$data, input$sum_vars, input$sum_name)
    
    tryCatch({
      if (input$sum_type == "row") {
        # Calculate row sums
        rv$data <- rv$data %>%
          mutate(!!input$sum_name := rowSums(select(., all_of(input$sum_vars)), na.rm = TRUE))
        
      } else {
        # Calculate column sums
        col_sums <- colSums(select(rv$data, all_of(input$sum_vars)), na.rm = TRUE)
        new_row <- as.list(rep(NA, ncol(rv$data)))
        names(new_row) <- names(rv$data)
        
        for (var in input$sum_vars) {
          new_row[[var]] <- col_sums[[var]]
        }
        
        # Add the sum row
        rv$data <- rv$data %>%
          add_row(!!!new_row) %>%
          mutate(across(where(is.character), ~replace_na(., "Column Sum")))
        
        # Set the sum name in the last row
        rv$data[nrow(rv$data), input$sum_name] <- "Column Sum"
      }
      
      showNotification("Sums calculated successfully!", type = "message")
    }, error = function(e) {
      showNotification(paste("Error calculating sums:", e$message), type = "error")
    })
  })
  
    # Create dummy variables
    observeEvent(input$apply_dummy, {
      req(rv$data, input$dummy_col)
      
      tryCatch({
        # Create dummy variables using fastDummies
        dummy_data <- dummy_cols(rv$data[[input$dummy_col]], 
                                 remove_first_dummy = FALSE,
                                 remove_selected_columns = FALSE)
        
        # Remove the original column if requested
        if (!input$keep_dummy_col) {
          rv$data <- rv$data %>%
            select(-all_of(input$dummy_col))
        }
        
        # Add the dummy variables to the data
        dummy_col_names <- names(dummy_data)
        for (col in dummy_col_names) {
          rv$data[[col]] <- dummy_data[[col]]
        }
        
        # Update last_saved version
        rv$last_saved <- rv$data
        showNotification("Dummy variables created successfully!", type = "message")
      }, error = function(e) {
        showNotification(paste("Error creating dummy variables:", e$message), type = "error")
      })
    })
    
    # Apply column reordering
    observeEvent(input$apply_col_order, {
      req(rv$data, input$col_order)
      
      tryCatch({
        # Get all columns including those not selected
        all_cols <- names(rv$data)
        selected_cols <- input$col_order
        remaining_cols <- setdiff(all_cols, selected_cols)
        
        # Reorder columns with selected ones first
        rv$data <- rv$data %>%
          select(all_of(selected_cols), all_of(remaining_cols))
        
        # Update last_saved version
        rv$last_saved <- rv$data
        showNotification("Columns reordered successfully!", type = "message")
      }, error = function(e) {
        showNotification(paste("Error reordering columns:", e$message), type = "error")
      })
    })
    
    # Dynamic UI for row operations
    output$row_op_ui <- renderUI({
      req(input$row_op, rv$data)
      
      switch(
        input$row_op,
        
        "Filter" = tagList(
          selectInput(
            "filter_col", 
            "Select Column", 
            choices = names(rv$data),
            width = "100%"
          ),
          selectInput(
            "filter_op", 
            "Operator",
            choices = c("==", "!=", ">", "<", ">=", "<=", "contains", "starts with", "ends with"),
            width = "100%"
          ),
          uiOutput("filter_val_ui"),
          actionBttn(
            "apply_filter",
            "Apply Filter",
            style = "material-flat",
            color = "primary",
            size = "sm"
          ),
          actionBttn(
            "reset_filter",
            "Reset Filter",
            style = "material-flat",
            color = "warning",
            size = "sm"
          )
        ),
        
        "Sort" = tagList(
          pickerInput(
            "sort_cols", 
            "Select Columns",
            choices = names(rv$data),
            multiple = TRUE,
            options = list(
              `actions-box` = TRUE,
              `selected-text-format` = "count > 2"
            )
          ),
          radioGroupButtons(
            "sort_order", 
            "Order",
            choices = c("Ascending" = "asc", "Descending" = "desc"),
            status = "primary"
          ),
          actionBttn(
            "apply_sort",
            "Apply Sort",
            style = "material-flat",
            color = "primary",
            size = "sm"
          )
        ),
        
        "Remove Duplicates" = tagList(
          radioGroupButtons(
            "dup_type",
            "Duplicate Type",
            choices = c("Entire Row" = "row", "Selected Columns" = "columns"),
            selected = "row",
            status = "primary"
          ),
          conditionalPanel(
            condition = "input.dup_type == 'columns'",
            pickerInput(
              "dup_cols", 
              "Columns to Check",
              choices = names(rv$data),
              multiple = TRUE,
              options = list(
                `actions-box` = TRUE,
                `selected-text-format` = "count > 2"
              )
            )
          ),
          actionBttn(
            "apply_dedup",
            "Remove Duplicates",
            style = "material-flat",
            color = "primary",
            size = "sm"
          )
        ),
        
        "Sample" = tagList(
          sliderInput(
            "sample_size", 
            "Number of Rows",
            min = 1, 
            max = ifelse(!is.null(rv$data), nrow(rv$data), 100),
            value = min(10, ifelse(!is.null(rv$data), nrow(rv$data), 10)),
            step = 1
          ),
          actionBttn(
            "apply_sample",
            "Take Sample",
            style = "material-flat",
            color = "primary",
            size = "sm"
          )
        )
      )
    })
    
  
  # Dynamic UI for filter value based on column type
  output$filter_val_ui <- renderUI({
    req(input$filter_col, rv$data)
    
    col <- rv$data[[input$filter_col]]
    
    if (is.numeric(col)) {
      numericInput("filter_val", "Value", value = 0)
    } else if (is.logical(col)) {
      selectInput("filter_val", "Value", choices = c(TRUE, FALSE))
    } else if (is.factor(col) || is.character(col)) {
      textInput("filter_val", "Value", value = "")
    } else if (inherits(col, "Date")) {
      dateInput("filter_val", "Value")
    } else {
      textInput("filter_val", "Value", value = "")
    }
  })
  
  # Apply header row selection
  observeEvent(input$apply_header, {
    req(input$file, input$header_row)
    
    tryCatch({
      ext <- tools::file_ext(input$file$name)
      path <- input$file$datapath
      
      # Read the file without header first
      raw_data <- switch(
        ext,
        csv = read_csv(
          path,
          col_names = FALSE,
          locale = locale(decimal_mark = ifelse(is.null(input$decimal), ".", input$decimal))
        ),
        sav = read_sav(path),
        dta = read_dta(path),
        xlsx = read_excel(path, col_names = FALSE),
        xls = read_excel(path, col_names = FALSE),
        validate("Invalid file; Please upload a .csv, .sav, .dta, .xlsx or .xls file")
      )
      
      # Get the specified header row
      header_row <- input$header_row
      
      # Validate header row
      if (header_row < 1 || header_row > nrow(raw_data)) {
        showNotification("Invalid header row number", type = "error")
        return()
      }
      
      # Set column names from the specified row
      col_names <- as.character(raw_data[header_row, ])
      
      # Remove all rows up to and including the header row
      if (nrow(raw_data) > header_row) {
        rv$data <- raw_data[-seq_len(header_row), ]
      } else {
        rv$data <- raw_data[0, ]  # Empty data frame with correct columns
      }
      
      # Set column names
      names(rv$data) <- col_names
      
      # Update original and last_saved
      rv$original <- rv$data
      rv$last_saved <- rv$data
      
      showNotification("Header applied successfully!", type = "message")
    }, error = function(e) {
      showNotification(paste("Error applying header:", e$message), type = "error")
    })
  })
  
  # Read data from uploaded file with header row selection
  observeEvent(input$file, {
    req(input$file)
    
    tryCatch({
      ext <- tools::file_ext(input$file$name)
      path <- input$file$datapath
      
      if (input$header_option == "first_row") {
        # Use first row as header
        rv$data <- switch(
          ext,
          csv = read_csv(
            path,
            col_names = TRUE,
            locale = locale(decimal_mark = ifelse(is.null(input$decimal), ".", input$decimal))
          ),
          sav = read_sav(path),
          dta = read_dta(path),
          xlsx = read_excel(path, col_names = TRUE),
          xls = read_excel(path, col_names = TRUE),
          validate("Invalid file; Please upload a .csv, .sav, .dta, .xlsx or .xls file")
        )
      } else if (input$header_option == "specific_row") {
        # This will be handled by the apply_header action button
        return()
      } else {
        # No header
        rv$data <- switch(
          ext,
          csv = read_csv(
            path,
            col_names = FALSE,
            locale = locale(decimal_mark = ifelse(is.null(input$decimal), ".", input$decimal))
          ),
          sav = read_sav(path),
          dta = read_dta(path),
          xlsx = read_excel(path, col_names = FALSE),
          xls = read_excel(path, col_names = FALSE),
          validate("Invalid file; Please upload a .csv, .sav, .dta, .xlsx or .xls file")
        )
      }
      
      rv$original <- rv$data
      rv$last_saved <- rv$data  # Initialize last_saved with original
      rv$edited_cells <- NULL
      showNotification("Data loaded successfully!", type = "message")
      
    }, error = function(e) {
      showNotification(paste("Error loading file:", e$message), type = "error")
    })
  })
  
  # Transpose data
  observeEvent(input$transpose, {
    req(rv$data)
    
    tryCatch({
      rv$data <- rv$data %>% 
        t() %>% 
        as_tibble(rownames = "Column") %>% 
        mutate(across(where(is.character), ~na_if(., "NA")))
      
      showNotification("Data transposed!", type = "message")
    }, error = function(e) {
      showNotification(paste("Error transposing data:", e$message), type = "error")
    })
  })
  
  # Reset data
  observeEvent(input$reset, {
    req(rv$original)
    rv$data <- rv$original
    rv$last_saved <- rv$original
    rv$edited_cells <- NULL
    showNotification("Data completely reset to original uploaded version", type = "message")
  })
  
  # Apply column operations
  observeEvent(input$apply_rename, {
    req(rv$data, input$rename_col, input$new_name)
    
    tryCatch({
      rv$data <- rv$data %>% 
        rename(!!input$new_name := !!input$rename_col)
      
      showNotification("Column renamed successfully!", type = "message")
    }, error = function(e) {
      showNotification(paste("Error renaming column:", e$message), type = "error")
    })
  })
  
  observeEvent(input$apply_recode, {
    req(rv$data, input$recode_col, input$recode_map)
    
    tryCatch({
      # Check if the column is numeric for range recoding
      col_is_numeric <- is.numeric(rv$data[[input$recode_col]])
      
      # Process recoding rules
      rules <- str_split(input$recode_map, ",\\s*")[[1]] %>% 
        str_trim() %>% 
        keep(~str_detect(., ".+=.+"))
      
    
    if (col_is_numeric) {
      # Process range recoding for numerical variables
      recode_exprs <- vector("list", length(rules))
      
      for (i in seq_along(rules)) {
        parts <- str_split(rules[i], "=", n = 2)[[1]] %>% str_trim()
        condition_part <- parts[1]
        label <- parts[2]
        
        if (str_detect(condition_part, "\\d+\\s*-\\s*\\d+")) {
          # Handle range (e.g., "1-5=Low")
          range_vals <- str_split(condition_part, "-")[[1]] %>% 
            str_trim() %>% 
            as.numeric()
          
          # Create safer evaluation environment
          recode_exprs[[i]] <- substitute(
            dplyr::between(x, min_val, max_val) ~ label,
            list(
              x = as.name(input$recode_col),
              min_val = range_vals[1],
              max_val = range_vals[2],
              label = label
            )
          )
        } else if (str_detect(condition_part, "^[<>=!]+\\d+")) {
          # Handle operator-based range (e.g., "<5=Low")
          op <- str_extract(condition_part, "^[<>=!]+")
          val <- str_extract(condition_part, "\\d+") %>% as.numeric()
          
          # Create safer evaluation environment
          recode_exprs[[i]] <- substitute(
            eval(parse(text = paste(x, op, val))) ~ label,
            list(
              x = as.name(input$recode_col),
              op = op,
              val = val,
              label = label
            )
          )
        } else {
          # Handle exact value
          val <- as.numeric(condition_part)
          recode_exprs[[i]] <- substitute(
            x == val ~ label,
            list(
              x = as.name(input$recode_col),
              val = val,
              label = label
            )
          )
        }
      }
      
      # Apply the recoding using case_when with explicit environment
      recoded_values <- rv$data %>%
        mutate(new_col = case_when(!!!unlist(recode_exprs))) %>%
        pull(new_col)
      
      # Replace NA values with original values
      recoded_values <- ifelse(is.na(recoded_values), 
                               as.character(rv$data[[input$recode_col]]), 
                               recoded_values)
      
    } else {
      # Process exact value recoding for categorical variables
      recode_pairs <- rules %>% 
        str_split_fixed("=", 2) %>% 
        as.data.frame() %>% 
        setNames(c("old", "new")) %>% 
        mutate_all(str_trim)
      
      recoded_values <- rv$data[[input$recode_col]] %>% 
        as.character() %>% 
        recode(!!!setNames(recode_pairs$new, recode_pairs$old))
    }
    
    # Determine where to store the recoded values
    if (input$recode_type == "new" && !is.null(input$new_recode_name)) {
      # Create new column
      rv$data[[input$new_recode_name]] <- recoded_values
    } else {
      # Replace existing column
      rv$data[[input$recode_col]] <- recoded_values
    }
    
    # Update last_saved version
    rv$last_saved <- rv$data
    showNotification("Recoding applied successfully!", type = "message")
    }, error = function(e) {
      showNotification(paste("Error recoding column:", e$message), type = "error")
    })
  })

observeEvent(input$apply_create, {
  req(rv$data, input$new_col_name)
  
  tryCatch({
    if (input$new_col_type == "Numeric") {
      rv$data[[input$new_col_name]] <- as.numeric(input$num_val)
    } else if (input$new_col_type == "Character") {
      rv$data[[input$new_col_name]] <- as.character(input$char_val)
    } else if (input$new_col_type == "Logical") {
      rv$data[[input$new_col_name]] <- FALSE
    } else if (input$new_col_type == "Date") {
      rv$data[[input$new_col_name]] <- as.Date(NA)
    }
    
    showNotification("Column created successfully!", type = "message")
  }, error = function(e) {
    showNotification(paste("Error creating column:", e$message), type = "error")
  })
})

observeEvent(input$apply_delete, {
  req(rv$data, input$delete_cols)
  
  tryCatch({
    rv$data <- rv$data %>% 
      select(-any_of(input$delete_cols))
    
    showNotification("Columns deleted successfully!", type = "message")
  }, error = function(e) {
    showNotification(paste("Error deleting columns:", e$message), type = "error")
  })
})

observeEvent(input$apply_convert, {
  req(rv$data, input$convert_col, input$new_type)
  
  tryCatch({
    rv$data <- rv$data %>% 
      mutate(!!input$convert_col := switch(
        input$new_type,
        "Numeric" = as.numeric(!!sym(input$convert_col)),
        "Character" = as.character(!!sym(input$convert_col)),
        "Factor" = as.factor(!!sym(input$convert_col)),
        "Date" = as.Date(!!sym(input$convert_col))
      ))
    
    showNotification("Column type converted successfully!", type = "message")
  }, error = function(e) {
    showNotification(paste("Error converting column type:", e$message), type = "error")
  })
})

# Apply row operations
observeEvent(input$apply_filter, {
  req(rv$data, input$filter_col, input$filter_op, input$filter_val)
  
  tryCatch({
    col <- sym(input$filter_col)
    val <- input$filter_val
    
    if (input$filter_op %in% c("contains", "starts with", "ends with")) {
      pattern <- switch(
        input$filter_op,
        "contains" = val,
        "starts with" = paste0("^", val),
        "ends with" = paste0(val, "$")
      )
      
      rv$data <- rv$data %>% 
        filter(str_detect(as.character(!!col), regex(pattern, ignore_case = TRUE)))
    } else {
      # For numerical/date comparisons, ensure the value is of the correct type
      if (is.numeric(rv$data[[input$filter_col]])) {
        val <- as.numeric(val)
      } else if (inherits(rv$data[[input$filter_col]], "Date")) {
        val <- as.Date(val)
      } else if (is.logical(rv$data[[input$filter_col]])) {
        val <- as.logical(val)
      }
      
      # Build the filter expression
      filter_expr <- switch(
        input$filter_op,
        "==" = expr(!!col == !!val),
        "!=" = expr(!!col != !!val),
        ">" = expr(!!col > !!val),
        "<" = expr(!!col < !!val),
        ">=" = expr(!!col >= !!val),
        "<=" = expr(!!col <= !!val)
      )
      
      rv$data <- rv$data %>% 
        filter(!!filter_expr)
    }
    
    showNotification("Filter applied successfully!", type = "message")
  }, error = function(e) {
    showNotification(paste("Error applying filter:", e$message), type = "error")
  })
})

# Reset filter - now reverts to last_saved version
observeEvent(input$reset_filter, {
  req(rv$last_saved)
  rv$data <- rv$last_saved
  showNotification("Filter reset successfully! Reverted to last saved version.", type = "message")
})

# Apply sort
observeEvent(input$apply_sort, {
  req(rv$data, input$sort_cols)
  
  tryCatch({
    if (input$sort_order == "desc") {
      rv$data <- rv$data %>%
        arrange(across(all_of(input$sort_cols), ~desc(.)))
    } else {
      rv$data <- rv$data %>%
        arrange(across(all_of(input$sort_cols)))
    }
    
    showNotification("Data sorted successfully!", type = "message")
  }, error = function(e) {
    showNotification(paste("Error sorting data:", e$message), type = "error")
  })
})

# Remove duplicates - FIXED
observeEvent(input$apply_dedup, {
  req(rv$data)
  
  tryCatch({
    if (input$dup_type == "columns" && length(input$dup_cols) > 0) {
      # Remove duplicates based on selected columns
      rv$data <- rv$data %>% 
        distinct(across(all_of(input$dup_cols)), .keep_all = TRUE)
    } else {
      # Remove duplicates based on entire row
      rv$data <- rv$data %>% 
        distinct()
    }
    
    # Update last_saved version
    rv$last_saved <- rv$data
    showNotification("Duplicates removed successfully!", type = "message")
  }, error = function(e) {
    showNotification(paste("Error removing duplicates:", e$message), type = "error")
  })
})

observeEvent(input$apply_sample, {
  req(rv$data, input$sample_size)
  
  tryCatch({
    n <- min(input$sample_size, nrow(rv$data))
    rv$data <- rv$data %>% 
      slice_sample(n = n)
    
    showNotification(paste("Sample of", n, "rows taken successfully!"), type = "message")
  }, error = function(e) {
    showNotification(paste("Error taking sample:", e$message), type = "error")
  })
})

# Handle cell editing
observeEvent(input$data_table_cell_edit, {
  info <- input$data_table_cell_edit
  i <- info$row
  j <- info$col + 1  # DT columns are 0-based
  v <- info$value
  
  # Store the edit
  rv$edited_cells <- rbind(rv$edited_cells, data.frame(row = i, col = j, value = v))
})

# Save cell edits
observeEvent(input$save_edits, {
  req(rv$edited_cells)
  
  tryCatch({
    for (k in 1:nrow(rv$edited_cells)) {
      i <- rv$edited_cells$row[k]
      j <- rv$edited_cells$col[k]
      v <- rv$edited_cells$value[k]
      
      # Convert value to appropriate type
      col_type <- class(rv$data[[j]])
      if (col_type == "numeric") {
        v <- as.numeric(v)
      } else if (col_type == "logical") {
        v <- as.logical(v)
      } else if (col_type == "factor") {
        v <- as.character(v)
      }
      
      rv$data[i, j] <- v
    }
    
    # Update last_saved version when edits are saved
    rv$last_saved <- rv$data
    rv$edited_cells <- NULL
    showNotification("Cell edits saved successfully!", type = "message")
  }, error = function(e) {
    showNotification(paste("Error saving edits:", e$message), type = "error")
  })
})

# Highlight missing data - UPDATED to handle categorical variables
observeEvent(input$show_missing, {
  req(input$missing_cols, rv$data)
  
  output$data_table <- renderDT({
    datatable(
      rv$data,
      extensions = c('Buttons', 'Scroller'),
      options = list(
        scrollX = TRUE,
        scrollY = 500,
        scroller = TRUE,
        dom = 'Bfrtip',
        buttons = c('copy', 'csv', 'excel', 'pdf', 'print'),
        pageLength = 10
      ),
      filter = 'top',
      rownames = FALSE,
      class = 'cell-border stripe hover',
      editable = TRUE  # Maintain editability
    ) %>% 
      formatStyle(
        input$missing_cols,
        backgroundColor = styleEqual(
          c(NA, ""),  # Highlight both NA and empty strings
          c("#ffcccc", "#ffcccc")  # Same color for both types of missing
        )
      )
  })
  
  # Show missing values summary for all variables
  output$missing_stats <- renderPrint({
    req(rv$data)
    
    cat("Missing Values Summary:\n\n")
    missing_counts <- rv$data %>% 
      summarise(across(everything(), ~sum(is.na(.) | . == ""))) %>% 
      pivot_longer(everything(), names_to = "Variable", values_to = "Missing")
    
    print(missing_counts)
    cat("\nTotal missing values:", sum(missing_counts$Missing))
  })
})

# Detect outliers
observeEvent(input$detect_outliers, {
  req(input$outlier_col, rv$data)
  
  col_data <- rv$data[[input$outlier_col]]
  mean_val <- mean(col_data, na.rm = TRUE)
  sd_val <- sd(col_data, na.rm = TRUE)
  threshold <- input$outlier_sd * sd_val
  
  outliers <- which(abs(col_data - mean_val) > threshold)
  
  if (length(outliers) > 0) {
    showModal(modalDialog(
      title = "Outliers Detected",
      sprintf("Found %d outliers in column '%s'", length(outliers), input$outlier_col),
      easyClose = TRUE,
      footer = NULL
    ))
    
    output$data_table <- renderDT({
      datatable(
        rv$data,
        extensions = c('Buttons', 'Scroller'),
        options = list(
          scrollX = TRUE,
          scrollY = 500,
          scroller = TRUE,
          dom = 'Bfrtip',
          buttons = c('copy', 'csv', 'excel', 'pdf', 'print'),
          pageLength = 10
        ),
        filter = 'top',
        rownames = FALSE,
        class = 'cell-border stripe hover',
        editable = TRUE  # Maintain editability
      ) %>% 
        formatStyle(
          input$outlier_col,
          backgroundColor = styleInterval(
            c(mean_val - threshold, mean_val + threshold),
            c("#ffcccc", "white", "#ffcccc")
          )
        )
    })
  } else {
    showNotification("No outliers detected!", type = "message")
  }
})

# Remove outliers
observeEvent(input$remove_outliers, {
  req(input$outlier_col, rv$data)
  
  col_data <- rv$data[[input$outlier_col]]
  mean_val <- mean(col_data, na.rm = TRUE)
  sd_val <- sd(col_data, na.rm = TRUE)
  threshold <- input$outlier_sd * sd_val
  
  rv$data <- rv$data %>% 
    filter(abs(!!sym(input$outlier_col) - mean_val) <= threshold)
  
  # Update last_saved version
  rv$last_saved <- rv$data
  showNotification(sprintf("Removed outliers from column '%s'", input$outlier_col), 
                   type = "message")
})

# Apply imputation - FIXED
observeEvent(input$apply_impute, {
  req(input$impute_cols, rv$data)
  
  tryCatch({
    for (col in input$impute_cols) {
      if (input$impute_method == "Mean") {
        impute_val <- mean(rv$data[[col]], na.rm = TRUE)
        rv$data <- rv$data %>%
          mutate(!!col := ifelse(is.na(!!sym(col)), impute_val, !!sym(col)))
      } else if (input$impute_method == "Median") {
        impute_val <- median(rv$data[[col]], na.rm = TRUE)
        rv$data <- rv$data %>%
          mutate(!!col := ifelse(is.na(!!sym(col)), impute_val, !!sym(col)))
      } else if (input$impute_method == "Mode") {
        impute_val <- as.numeric(names(sort(table(rv$data[[col]]), decreasing = TRUE))[1])
        rv$data <- rv$data %>%
          mutate(!!col := ifelse(is.na(!!sym(col)), impute_val, !!sym(col)))
      } else if (input$impute_method == "Zero") {
        rv$data <- rv$data %>%
          mutate(!!col := ifelse(is.na(!!sym(col)), 0, !!sym(col)))
      } else if (input$impute_method == "Linear Interpolation") {
        # Use zoo's na.approx for interpolation
        rv$data <- rv$data %>%
          mutate(!!col := zoo::na.approx(!!sym(col), na.rm = FALSE, rule = 2))
      } else if (input$impute_method == "KNN") {
        # Use VIM's kNN for better KNN imputation
        temp_data <- rv$data %>%
          select(where(is.numeric)) %>%
          as.data.frame()
        
        # Check if column exists in numeric data
        if (!col %in% names(temp_data)) {
          showNotification(paste("Column", col, "is not numeric and cannot be used with KNN"), 
                           type = "error")
          next
        }
        
        # Impute using VIM's kNN
        imputed_data <- VIM::kNN(temp_data, k = input$k_value, imp_var = FALSE)
        
        # Update the column in the original data
        rv$data[[col]] <- imputed_data[[col]]
      }
    }
    
    # Update last_saved version
    rv$last_saved <- rv$data
    showNotification("Imputation applied successfully!", type = "message")
  }, error = function(e) {
    showNotification(paste("Error applying imputation:", e$message), type = "error")
  })
})

# Run descriptive statistics - IMPROVED to handle categorical missing values
observeEvent(input$run_desc, {
  req(rv$data, input$desc_var)
  
  output$descriptive_stats <- renderPrint({
    var <- rv$data[[input$desc_var]]
    
    if (is.numeric(var)) {
      cat("Numeric Variable Summary for:", input$desc_var, "\n\n")
      cat("Mean:", mean(var, na.rm = TRUE), "\n")
      cat("Median:", median(var, na.rm = TRUE), "\n")
      cat("Mode:", DescTools::Mode(var, na.rm = TRUE), "\n")
      cat("Standard Deviation:", sd(var, na.rm = TRUE), "\n")
      cat("Variance:", var(var, na.rm = TRUE), "\n")
      cat("Minimum:", min(var, na.rm = TRUE), "\n")
      cat("Maximum:", max(var, na.rm = TRUE), "\n")
      cat("Range:", diff(range(var, na.rm = TRUE)), "\n")
      cat("Skewness:", DescTools::Skew(var, na.rm = TRUE), "\n")
      cat("Kurtosis:", DescTools::Kurt(var, na.rm = TRUE), "\n")
      cat("Number of Missing Values:", sum(is.na(var) | var == ""), "\n")
      cat("Number of Valid Values:", sum(!is.na(var) & var != ""), "\n")
      
      cat("\nQuantiles:\n")
      print(quantile(var, probs = seq(0, 1, 0.1), na.rm = TRUE))
      
    } else {
      cat("Categorical Variable Summary for:", input$desc_var, "\n\n")
      freq_table <- table(var, useNA = "ifany")
      prop_table <- prop.table(freq_table) * 100
      
      cat("Frequency Table (including missing values):\n")
      print(freq_table)
      cat("\nPercentage Table:\n")
      print(round(prop_table, 1))
      cat("\nMode:", DescTools::Mode(var, na.rm = TRUE), "\n")
      cat("Number of Categories:", length(unique(var)), "\n")
      cat("Number of Missing Values:", sum(is.na(var) | var == ""), "\n")
    }
  })
})

# Missing data plot
output$missing_plot <- renderPlotly({
  req(rv$data)
  
  missing_data <- rv$data %>% 
    summarise(across(everything(), ~sum(is.na(.) | . == ""))) %>% 
    pivot_longer(everything(), names_to = "Variable", values_to = "Missing") %>% 
    mutate(Percentage = Missing / nrow(rv$data) * 100)
  
  plot_ly(missing_data, x = ~Variable, y = ~Percentage, type = 'bar',
          marker = list(color = '#ff6666'),
          hoverinfo = 'text',
          text = ~paste("Variable:", Variable, "<br>Missing:", Missing, "<br>Percentage:", round(Percentage, 1), "%")) %>% 
    layout(title = "Missing Data by Variable",
           xaxis = list(title = ""),
           yaxis = list(title = "Percentage Missing"))
})

# Missing data summary - UPDATED to include empty strings as missing
output$missing_summary <- renderPrint({
  req(rv$data)
  
  cat("Missing Data Summary (NA or empty strings):\n\n")
  print(colSums(is.na(rv$data) | rv$data == ""))
  cat("\n\nTotal missing values:", sum(is.na(rv$data) | rv$data == ""))
  cat("\nPercentage missing:", round(mean(is.na(rv$data) | rv$data == "") * 100, 1), "%\n")
})

# Data summary
output$summary <- renderPrint({
  req(rv$data)
  summary(rv$data)
})

# Data structure
output$structure <- renderPrint({
  req(rv$data)
  glimpse(rv$data)
})

# Data plot
output$data_plot <- renderPlot({
  req(rv$data, input$plot_type)
  
  if (input$plot_type == "Histogram") {
    req(input$hist_col)
    ggplot(rv$data, aes(x = !!sym(input$hist_col))) +
      geom_histogram(fill = "#18bc9c", color = "white", bins = 30) +
      labs(title = paste("Distribution of", input$hist_col)) +
      theme_minimal()
  } else if (input$plot_type == "Bar Plot") {
    req(input$bar_col)
    ggplot(rv$data, aes(x = !!sym(input$bar_col))) +
      geom_bar(fill = "#18bc9c", color = "white") +
      labs(title = paste("Frequency of", input$bar_col)) +
      theme_minimal() +
      theme(axis.text.x = element_text(angle = 45, hjust = 1))
  } else if (input$plot_type == "Box Plot") {
    req(input$box_col)
    if (input$box_group != "None") {
      ggplot(rv$data, aes(x = !!sym(input$box_group), y = !!sym(input$box_col), fill = !!sym(input$box_group))) +
        geom_boxplot() +
        labs(title = paste("Box plot of", input$box_col, "by", input$box_group)) +
        theme_minimal() +
        theme(legend.position = "none")
    } else {
      ggplot(rv$data, aes(y = !!sym(input$box_col))) +
        geom_boxplot(fill = "#18bc9c") +
        labs(title = paste("Box plot of", input$box_col)) +
        theme_minimal()
    }
  } else if (input$plot_type == "Scatter Plot") {
    req(input$scatter_x, input$scatter_y)
    if (input$scatter_color != "None") {
      ggplot(rv$data, aes(x = !!sym(input$scatter_x), y = !!sym(input$scatter_y), color = !!sym(input$scatter_color))) +
        geom_point() +
        labs(title = paste(input$scatter_y, "vs", input$scatter_x)) +
        theme_minimal()
    } else {
      ggplot(rv$data, aes(x = !!sym(input$scatter_x), y = !!sym(input$scatter_y))) +
        geom_point(color = "#18bc9c") +
        labs(title = paste(input$scatter_y, "vs", input$scatter_x)) +
        theme_minimal()
    }
  }
})

# Display data table with editing capability - FIXED to make all columns editable
output$data_table <- renderDT({
  req(rv$data)
  
  datatable(
    rv$data,
    extensions = c('Buttons', 'Scroller'),
    options = list(
      scrollX = TRUE,
      scrollY = 500,
      scroller = TRUE,
      dom = 'Bfrtip',
      buttons = c('copy', 'csv', 'excel', 'pdf', 'print'),
      pageLength = 10
    ),
    filter = 'top',
    rownames = FALSE,
    editable = TRUE,  # Now all columns are editable
    class = 'cell-border stripe hover',
    selection = 'none'
  )
})

# Download processed data in selected format
output$downloadData <- downloadHandler(
  filename = function() {
    paste0("processed_data_", Sys.Date(), 
           switch(input$download_format,
                  "CSV" = ".csv",
                  "Excel" = ".xlsx",
                  "SPSS" = ".sav",
                  "Stata" = ".dta"))
  },
  content = function(file) {
    # Subset data if specific columns are selected
    if (!is.null(input$download_cols) && length(input$download_cols) > 0) {
      download_data <- rv$data %>% select(all_of(input$download_cols))
    } else {
      download_data <- rv$data
    }
    
    switch(input$download_format,
           "CSV" = write_csv(download_data, file),
           "Excel" = writexl::write_xlsx(download_data, file),
           "SPSS" = write_sav(download_data, file),
           "Stata" = write_dta(download_data, file))
  }
)
}

shinyApp(ui, server)
