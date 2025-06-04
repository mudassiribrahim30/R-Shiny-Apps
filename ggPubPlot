# Load required packages
library(shiny)
library(shinythemes)
library(tidyverse)
library(ggplot2)
library(readxl)
library(haven)
library(readr)
library(DT)
library(ggiraph)
library(patchwork)
library(ggrepel)
library(ggforce)
library(ggridges)
library(ggpubr)
library(ggtext)
library(bslib)
library(waffle)
library(ggwordcloud)
library(GGally)
library(ggalluvial)
library(ggsci)
library(colourpicker)
library(nortest) # For Anderson-Darling test

# Define UI
ui <- fluidPage(
  theme = bslib::bs_theme(
    bootswatch = "flatly",
    primary = "#3498db",
    secondary = "#2ecc71",
    success = "#2ecc71",
    info = "#3498db",
    warning = "#f39c12",
    danger = "#e74c3c",
    bg = "#f8f9fa",
    fg = "#212529",
    base_font = font_google("Open Sans")
  ),
  div(
    style = "background-color: #2c3e50; color: white; padding: 15px; margin-bottom: 20px; border-radius: 5px;",
    h2("ggPubPlot", style = "margin: 0; font-weight: bold;"),
    p("", style = "margin: 0; font-size: 16px; opacity: 0.9;")
  ),
  sidebarLayout(
    sidebarPanel(
      width = 3,
      style = "background-color: #f8f9fa; border-radius: 5px; padding: 15px; height: 90vh; overflow-y: auto;",
      h4("Data Input", style = "color: #3498db; border-bottom: 2px solid #3498db; padding-bottom: 5px;"),
      radioButtons(
        "data_source",
        "Data Source:",
        choices = c("Upload File" = "upload", "Sample Data" = "sample"),
        selected = "upload",
        inline = TRUE
      ),
      conditionalPanel(
        condition = "input.data_source == 'upload'",
        fileInput(
          "file_upload",
          "Upload Data File",
          accept = c(
            ".csv", ".xlsx", ".xls",
            ".sav", ".dta", ".rds", ".tsv"
          ),
          multiple = FALSE,
          buttonLabel = "Browse...",
          placeholder = "No file selected"
        ),
        selectInput(
          "file_type",
          "File Type:",
          choices = c(
            "Auto-detect" = "auto",
            "CSV" = "csv",
            "Excel" = "excel",
            "SPSS" = "spss",
            "Stata" = "stata",
            "TSV" = "tsv"
          ),
          selected = "auto"
        ),
        conditionalPanel(
          condition = "input.file_type == 'csv' || input.file_type == 'tsv'",
          checkboxInput("header", "Header", TRUE),
          radioButtons("sep", "Separator",
                       choices = c(
                         Comma = ",",
                         Semicolon = ";",
                         Tab = "\t",
                         Space = " "
                       ),
                       selected = ","
          ),
          radioButtons("quote", "Quote",
                       choices = c(
                         None = "",
                         "Double Quote" = "\"",
                         "Single Quote" = "'"
                       ),
                       selected = "\""
          )
        ),
        conditionalPanel(
          condition = "input.file_type == 'excel'",
          numericInput("sheet", "Sheet Number", value = 1, min = 1)
        )
      ),
      conditionalPanel(
        condition = "input.data_source == 'sample'",
        selectInput(
          "sample_data",
          "Sample Dataset:",
          choices = c(
            "mpg" = "mpg",
            "diamonds" = "diamonds",
            "iris" = "iris",
            "mtcars" = "mtcars",
            "gapminder" = "gapminder",
            "economics" = "economics",
            "txhousing" = "txhousing",
            "luv_colours" = "luv_colours"
          )
        )
      ),
      hr(),
      h4("Plot Configuration", style = "color: #3498db; border-bottom: 2px solid #3498db; padding-bottom: 5px;"),
      selectInput(
        "plot_type",
        "Plot Type:",
        choices = c(
          "Scatter Plot" = "scatter",
          "Line Plot" = "line",
          "Bar Plot" = "bar",
          "Histogram" = "histogram",
          "Density Plot" = "density",
          "Box Plot" = "boxplot",
          "Violin Plot" = "violin",
          "Area Plot" = "area",
          "Tile Plot" = "tile",
          "Raster Plot" = "raster",
          "Dot Plot" = "dotplot",
          "Smooth Line" = "smooth",
          "Quantile Regression" = "quantile",
          "Path Plot" = "path",
          "Ribbon Plot" = "ribbon",
          "Polygon Plot" = "polygon",
          "Contour Plot" = "contour",
          "2D Density" = "density2d",
          "Binned Heatmap" = "bin2d",
          "Hexbin Plot" = "hex",
          "Jitter Plot" = "jitter",
          "Point Range" = "pointrange",
          "Error Bars" = "errorbar",
          "Crossbar" = "crossbar",
          "Linerange" = "linerange",
          "Step Plot" = "step",
          "Map" = "map",
          "Pie Chart" = "pie",
          "Donut Chart" = "donut",
          "Waffle Chart" = "waffle",
          "Alluvial" = "alluvial",
          "Sina Plot" = "sina",
          "Ridgeline Plot" = "ridgeline",
          "Parallel Coordinates" = "parcoord",
          "Word Cloud" = "wordcloud",
          "Clustered Bar Chart" = "clustered_bar"  # Added clustered bar chart option
        ),
        selected = "scatter"
      ),
      conditionalPanel(
        condition = "input.plot_type != 'pie' && input.plot_type != 'donut' && input.plot_type != 'waffle' && input.plot_type != 'wordcloud'",
        selectInput("x_var", "X Variable:", choices = NULL)
      ),
      conditionalPanel(
        condition = "input.plot_type != 'histogram' && input.plot_type != 'density' && input.plot_type != 'density2d' && input.plot_type != 'bin2d' && input.plot_type != 'hex' && input.plot_type != 'pie' && input.plot_type != 'donut' && input.plot_type != 'waffle' && input.plot_type != 'wordcloud'",
        selectInput("y_var", "Y Variable:", choices = NULL)
      ),
      conditionalPanel(
        condition = "input.plot_type == 'pie' || input.plot_type == 'donut'",
        selectInput("pie_var", "Categorical Variable:", choices = NULL)
      ),
      conditionalPanel(
        condition = "input.plot_type == 'waffle'",
        selectInput("waffle_var", "Categorical Variable:", choices = NULL)
      ),
      conditionalPanel(
        condition = "input.plot_type == 'wordcloud'",
        selectInput("word_var", "Text Variable:", choices = NULL),
        selectInput("freq_var", "Frequency Variable:", choices = NULL)
      ),
      conditionalPanel(
        condition = "input.plot_type == 'clustered_bar'",
        selectInput("cluster_var", "Cluster Variable:", choices = NULL)
      ),
      selectInput("color_var", "Color Variable (Optional):", choices = c("None" = "none")),
      selectInput("size_var", "Size Variable (Optional):", choices = c("None" = "none")),
      selectInput("shape_var", "Shape Variable (Optional):", choices = c("None" = "none")),
      selectInput("facet_var", "Facet Variable (Optional):", choices = c("None" = "none")),
      selectInput("group_var", "Group Variable (Optional):", choices = c("None" = "none")),
      conditionalPanel(
        condition = "input.plot_type == 'scatter' || input.plot_type == 'boxplot' || input.plot_type == 'violin' || input.plot_type == 'jitter'",
        checkboxInput("show_ids", "Show Data Point IDs", FALSE),
        conditionalPanel(
          condition = "input.show_ids == true",
          selectInput("id_var", "ID Variable:", choices = NULL)
        )
      ),
      conditionalPanel(
        condition = "input.plot_type == 'boxplot'",
        checkboxInput("show_outliers", "Identify Outliers", FALSE)
      ),
      
      # Plot dimensions section
      hr(),
      h4("Plot Dimensions", style = "color: #3498db; border-bottom: 2px solid #3498db; padding-bottom: 5px;"),
      sliderInput("plot_width", "Plot Width (inches):", min = 4, max = 20, value = 10, step = 0.5),
      sliderInput("plot_height", "Plot Height (inches):", min = 4, max = 20, value = 6, step = 0.5),
      sliderInput("plot_dpi", "Resolution (DPI):", min = 72, max = 600, value = 300, step = 10),
      
      # Color customization section
      hr(),
      h4("Color Customization", style = "color: #3498db; border-bottom: 2px solid #3498db; padding-bottom: 5px;"),
      radioButtons("color_choice", "Color Options:",
                   choices = c("Single Color" = "single", "Custom by Category" = "category"),
                   selected = "single",
                   inline = TRUE),
      conditionalPanel(
        condition = "input.color_choice == 'single'",
        colourpicker::colourInput("single_color", "Select Color:", value = "#3498db", showColour = "both")
      ),
      conditionalPanel(
        condition = "input.color_choice == 'category' && (input.color_var != 'none' || input.plot_type == 'pie' || input.plot_type == 'donut' || input.plot_type == 'bar' || input.plot_type == 'clustered_bar')",
        div(
          style = "border: 1px solid #ddd; border-radius: 5px; padding: 10px; margin-bottom: 15px;",
          h5("Category Colors", style = "margin-top: 0;"),
          div(
            style = "resize: both; overflow: auto; max-height: 400px; min-height: 100px; border: 1px solid #eee; padding: 10px;",
            uiOutput("category_colors_ui")
          ),
          helpText("Tip: Click the color box, then use arrow keys to browse and select a color", style = "font-size: 12px; color: #666;")
        )
      ),
      
      # Bar chart specific options
      conditionalPanel(
        condition = "input.plot_type == 'bar' || input.plot_type == 'clustered_bar'",
        checkboxInput("show_values", "Show Values on Bars", TRUE),
        conditionalPanel(
          condition = "input.show_values",
          numericInput("value_size", "Value Size:", value = 4, min = 1, max = 10),
          numericInput("value_digits", "Decimal Digits:", value = 1, min = 0, max = 5),
          checkboxInput("show_percentage", "Show Percentage", FALSE),
          conditionalPanel(
            condition = "input.show_percentage",
            checkboxInput("show_both", "Show Both Count and Percentage", FALSE)
          )
        )
      ),
      
      # Pie chart specific options
      conditionalPanel(
        condition = "input.plot_type == 'pie' || input.plot_type == 'donut'",
        checkboxInput("pie_labels", "Show Labels", TRUE),
        conditionalPanel(
          condition = "input.pie_labels",
          numericInput("label_size", "Label Size:", value = 4, min = 1, max = 10),
          numericInput("label_digits", "Decimal Digits:", value = 1, min = 0, max = 5)
        )
      ),
      
      hr(),
      h4("Aesthetics & Options", style = "color: #3498db; border-bottom: 2px solid #3498db; padding-bottom: 5px;"),
      conditionalPanel(
        condition = "input.plot_type == 'scatter' || input.plot_type == 'line' || input.plot_type == 'pointrange' || input.plot_type == 'jitter'",
        sliderInput("point_size", "Point Size:", min = 0.1, max = 10, value = 2, step = 0.1),
        sliderInput("point_alpha", "Point Transparency:", min = 0, max = 1, value = 1, step = 0.05)
      ),
      conditionalPanel(
        condition = "input.plot_type == 'line' || input.plot_type == 'path' || input.plot_type == 'step'",
        sliderInput("line_size", "Line Size:", min = 0.1, max = 5, value = 1, step = 0.1),
        sliderInput("line_alpha", "Line Transparency:", min = 0, max = 1, value = 1, step = 0.05),
        selectInput(
          "line_type",
          "Line Type:",
          choices = c(
            "solid", "dashed", "dotted",
            "dotdash", "longdash", "twodash"
          ),
          selected = "solid"
        )
      ),
      conditionalPanel(
        condition = "input.plot_type == 'bar' || input.plot_type == 'histogram' || input.plot_type == 'density' || input.plot_type == 'area' || input.plot_type == 'violin' || input.plot_type == 'boxplot' || input.plot_type == 'ribbon' || input.plot_type == 'polygon' || input.plot_type == 'clustered_bar'",
        sliderInput("fill_alpha", "Fill Transparency:", min = 0, max = 1, value = 0.7, step = 0.05)
      ),
      conditionalPanel(
        condition = "input.plot_type == 'bar' || input.plot_type == 'histogram' || input.plot_type == 'boxplot' || input.plot_type == 'violin' || input.plot_type == 'dotplot' || input.plot_type == 'crossbar' || input.plot_type == 'linerange' || input.plot_type == 'pointrange' || input.plot_type == 'errorbar' || input.plot_type == 'clustered_bar'",
        radioButtons(
          "position",
          "Position Adjustment:",
          choices = c(
            "stack" = "stack",
            "dodge" = "dodge",
            "fill" = "fill",
            "identity" = "identity"
          ),
          selected = "dodge",
          inline = TRUE
        )
      ),
      conditionalPanel(
        condition = "input.plot_type == 'histogram' || input.plot_type == 'density'",
        sliderInput("binwidth", "Bin Width:", min = 0.1, max = 10, value = 1, step = 0.1),
        numericInput("bins", "Number of Bins:", value = 30, min = 1)
      ),
      conditionalPanel(
        condition = "input.plot_type == 'smooth'",
        selectInput(
          "smooth_method",
          "Smoothing Method:",
          choices = c(
            "loess" = "loess",
            "gam" = "gam",
            "lm" = "lm",
            "glm" = "glm",
            "rlm" = "rlm"
          ),
          selected = "loess"
        ),
        sliderInput("smooth_span", "Smooth Span:", min = 0.1, max = 1, value = 0.75, step = 0.05),
        checkboxInput("se", "Show Confidence Interval", TRUE)
      ),
      conditionalPanel(
        condition = "input.plot_type == 'boxplot' || input.plot_type == 'violin'",
        checkboxInput("notch", "Notched Boxplot", FALSE),
        sliderInput("width", "Width:", min = 0.1, max = 1, value = 0.7, step = 0.05),
        checkboxInput("rotate_plot", "Rotate Plot (Horizontal)", FALSE)
      ),
      conditionalPanel(
        condition = "input.plot_type == 'bar' || input.plot_type == 'clustered_bar'",
        checkboxInput("rotate_barplot", "Rotate Plot (Horizontal)", FALSE)
      ),
      conditionalPanel(
        condition = "input.plot_type == 'violin'",
        checkboxInput("trim", "Trim Violins", TRUE),
        sliderInput("violin_scale", "Violin Scale:", min = 0.1, max = 2, value = 1, step = 0.1)
      ),
      conditionalPanel(
        condition = "input.plot_type == 'density2d' || input.plot_type == 'contour'",
        sliderInput("contour_size", "Contour Size:", min = 0.1, max = 2, value = 0.5, step = 0.1),
        numericInput("contour_bins", "Number of Contour Levels:", value = 10, min = 1)
      ),
      conditionalPanel(
        condition = "input.plot_type == 'hex'",
        numericInput("hex_bins", "Number of Hex Bins:", value = 30, min = 1)
      ),
      conditionalPanel(
        condition = "input.plot_type == 'pie' || input.plot_type == 'donut'",
        sliderInput("pie_size", "Pie Size:", min = 0.1, max = 1, value = 1, step = 0.05)
      ),
      conditionalPanel(
        condition = "input.plot_type == 'waffle'",
        numericInput("waffle_rows", "Number of Rows:", value = 10, min = 1),
        numericInput("waffle_cols", "Number of Columns:", value = 10, min = 1)
      ),
      
      # Statistical tests section - Updated with action button
      hr(),
      h4("Statistical Tests", style = "color: #3498db; border-bottom: 2px solid #3498db; padding-bottom: 5px;"),
      p("Click the button below to run statistical tests and see the results.", 
        style = "font-size: 12px; color: #666; margin-bottom: 10px;"),
      
      # Action button to run statistical tests
      actionButton("run_tests", "Run Statistical Tests", 
                   class = "btn-primary",
                   style = "background-color: #3498db; border-color: #2980b9; margin-bottom: 15px;"),
      
      conditionalPanel(
        condition = "input.plot_type == 'scatter'",
        checkboxInput("show_correlation", "Show Correlation", FALSE),
        conditionalPanel(
          condition = "input.show_correlation == true",
          radioButtons("cor_method", "Correlation Method:",
                       choices = c("Pearson" = "pearson", "Spearman" = "spearman", "Kendall" = "kendall"),
                       selected = "pearson",
                       inline = TRUE)
        )
      ),
      conditionalPanel(
        condition = "input.plot_type == 'boxplot' || input.plot_type == 'violin'",
        checkboxInput("show_test", "Show Statistical Test", FALSE),
        conditionalPanel(
          condition = "input.show_test == true",
          selectInput("test_method", "Test Method:",
                      choices = c("Wilcoxon" = "wilcox.test", "t-test" = "t.test", "Kruskal-Wallis" = "kruskal.test"),
                      selected = "wilcox.test")
        )
      ),
      conditionalPanel(
        condition = "input.plot_type == 'bar' && input.x_var != 'none' && input.y_var != 'none'",
        checkboxInput("show_chisq", "Show Chi-square Test", FALSE)
      ),
      conditionalPanel(
        condition = "input.plot_type == 'histogram'",
        checkboxInput("show_normality", "Show Normality Tests", FALSE)
      ),
      
      hr(),
      h4("Labels & Titles", style = "color: #3498db; border-bottom: 2px solid #3498db; padding-bottom: 5px;"),
      textInput("title", "Plot Title:", value = ""),
      textInput("x_lab", "X Axis Label:", value = ""),
      textInput("y_lab", "Y Axis Label:", value = ""),
      textInput("caption", "Caption:", value = ""),
      textInput("tag", "Tag:", value = ""),
      hr(),
      h4("Theme & Appearance", style = "color: #3498db; border-bottom: 2px solid #3498db; padding-bottom: 5px;"),
      selectInput(
        "theme",
        "Theme:",
        choices = c(
          "Gray" = "theme_gray",
          "Classic" = "theme_classic",
          "Minimal" = "theme_minimal",
          "BW" = "theme_bw",
          "Light" = "theme_light",
          "Dark" = "theme_dark",
          "Void" = "theme_void",
          "Test" = "theme_test",
          "Pubr" = "theme_pubr",
          "Pubclean" = "theme_pubclean",
          "Linedraw" = "theme_linedraw"
        ),
        selected = "theme_gray"
      ),
      sliderInput("base_size", "Base Font Size:", min = 8, max = 24, value = 12, step = 1),
      checkboxInput("coord_flip", "Flip Coordinates", FALSE),
      checkboxInput("interactive", "Interactive Plot", FALSE),
      hr(),
      div(
        style = "display: flex; justify-content: space-between;",
        actionButton("run_analysis", "Generate Plot", 
                     class = "btn-primary",
                     style = "background-color: #2ecc71; border-color: #27ae60;"),
        downloadButton("download_plot", "Download Plot", 
                       class = "btn-success",
                       style = "background-color: #3498db; border-color: #2980b9;"),
        downloadButton("download_data", "Download Data", 
                       class = "btn-info",
                       style = "background-color: #9b59b6; border-color: #8e44ad;")
      ),
      hr(),
      div(
        style = "background-color: #f1f1f1; padding: 10px; border-radius: 5px; margin-top: 20px;",
        p(strong("Developer Info")),
        p("Name: Mudasir Mohammed Ibrahim"),
        p("For suggestions or problems, contact:"),
        p(HTML("<a href='mailto:mudassiribrahim30@gmail.com'>mudassiribrahim30@gmail.com</a>"))
      )
    ),
    mainPanel(
      width = 9,
      style = "background-color: white; border-radius: 5px; padding: 15px; height: 90vh; overflow-y: auto;",
      tabsetPanel(
        type = "tabs",
        tabPanel(
          "Plot",
          conditionalPanel(
            condition = "input.interactive == false",
            plotOutput("static_plot", height = "600px")
          ),
          conditionalPanel(
            condition = "input.interactive == true",
            girafeOutput("interactive_plot", height = "600px")
          ),
          # Statistical test results section
          conditionalPanel(
            condition = "input.run_tests > 0",
            h4("Statistical Test Results"),
            verbatimTextOutput("statistical_results")
          )
        ),
        tabPanel(
          "Data",
          DTOutput("data_table")
        ),
        tabPanel(
          "Usage",
          verbatimTextOutput("plot_code"),
          actionButton("copy_code", "Thank You for Using the App", 
                       class = "btn-primary",
                       style = "background-color: #3498db; border-color: #2980b9; margin-top: 10px;")
        )
      )
    )
  )
)

# Define server logic
server <- function(input, output, session) {
  # Reactive data
  data <- reactive({
    if (input$data_source == "sample") {
      switch(input$sample_data,
             "mpg" = ggplot2::mpg,
             "diamonds" = ggplot2::diamonds,
             "iris" = datasets::iris,
             "mtcars" = datasets::mtcars |> tibble::rownames_to_column("model"),
             "gapminder" = {
               if (!requireNamespace("gapminder", quietly = TRUE)) {
                 stop("Please install the gapminder package: install.packages('gapminder')")
               }
               gapminder::gapminder
             },
             "economics" = ggplot2::economics,
             "txhousing" = ggplot2::txhousing,
             "luv_colours" = grDevices::luv_colours
      )
    } else {
      req(input$file_upload)
      ext <- tools::file_ext(input$file_upload$name)
      
      if (input$file_type == "auto") {
        file_type <- ext
      } else {
        file_type <- input$file_type
      }
      
      tryCatch({
        switch(file_type,
               "csv" = readr::read_csv(
                 input$file_upload$datapath,
                 col_names = input$header,
                 quote = input$quote
               ),
               "tsv" = readr::read_tsv(
                 input$file_upload$datapath,
                 col_names = input$header,
                 quote = input$quote
               ),
               "xlsx" = readxl::read_excel(
                 input$file_upload$datapath,
                 sheet = input$sheet
               ),
               "xls" = readxl::read_excel(
                 input$file_upload$datapath,
                 sheet = input$sheet
               ),
               "sav" = haven::read_spss(input$file_upload$datapath),
               "dta" = haven::read_stata(input$file_upload$datapath),
               "rds" = readRDS(input$file_upload$datapath),
               stop("Unsupported file type")
        )
      }, error = function(e) {
        showNotification(paste("Error reading file:", e$message), type = "error")
        NULL
      })
    }
  })
  
  # Generate UI for category colors when a color variable is selected
  output$category_colors_ui <- renderUI({
    # Determine which variable to use for categories
    var_to_use <- if (input$plot_type %in% c("pie", "donut")) {
      input$pie_var
    } else if (input$plot_type == "bar" && input$color_var == "none") {
      input$x_var
    } else if (input$plot_type == "clustered_bar") {
      input$cluster_var
    } else {
      input$color_var
    }
    
    if (input$color_choice == "category" && var_to_use != "none") {
      df <- data()
      if (!is.null(df)) {
        if (var_to_use %in% names(df)) {
          # Get unique categories, remove NA values, and sort them
          categories <- sort(unique(na.omit(df[[var_to_use]])))
          
          # Create a scrollable container for the color pickers
          div(
            style = "max-height: 300px; overflow-y: auto; border: 1px solid #ddd; padding: 10px; border-radius: 5px;",
            h5(paste("Custom Colors for", var_to_use, "Categories:"), style = "margin-top: 0;"),
            lapply(1:length(categories), function(i) {
              fluidRow(
                column(
                  6,
                  style = "display: flex; align-items: center;",
                  div(
                    style = "margin-right: 10px; min-width: 100px;",
                    tags$span(categories[i])
                  )
                ),
                column(
                  6,
                  colourInput(
                    inputId = paste0("color_", i),
                    label = NULL,
                    value = scales::hue_pal()(length(categories))[i],
                    showColour = "background",
                    allowTransparent = FALSE
                  )
                )
              )
            })
          )
        }
      }
    }
  })
  
  # Get category colors
  category_colors <- reactive({
    # Determine which variable to use for categories
    var_to_use <- if (input$plot_type %in% c("pie", "donut")) {
      input$pie_var
    } else if (input$plot_type == "bar" && input$color_var == "none") {
      input$x_var
    } else if (input$plot_type == "clustered_bar") {
      input$cluster_var
    } else {
      input$color_var
    }
    
    if (input$color_choice == "category" && var_to_use != "none") {
      df <- data()
      if (!is.null(df)) {
        if (var_to_use %in% names(df)) {
          # Get unique categories, remove NA values, and sort them
          categories <- sort(unique(na.omit(df[[var_to_use]])))
          
          # Collect all the color inputs
          colors <- lapply(1:length(categories), function(i) {
            input[[paste0("color_", i)]]
          })
          
          # Filter out NULL values and create named vector
          valid_colors <- !sapply(colors, is.null)
          colors <- unlist(colors[valid_colors])
          categories <- categories[valid_colors]
          
          if (length(colors) > 0) {
            names(colors) <- categories
            return(colors)
          }
        }
      }
    }
    return(NULL)
  })
  
  # Update variable selections based on data
  observe({
    df <- data()
    if (!is.null(df)) {
      # Update x_var with "None" option for plots that can work without y_var
      x_choices <- c(names(df))
      if (input$plot_type %in% c("bar", "histogram", "density", "clustered_bar")) {
        x_choices <- c("None" = "none", names(df))
      }
      updateSelectInput(session, "x_var", choices = x_choices)
      
      # Update y_var with "None" option for plots that can work without y_var
      y_choices <- c(names(df))
      if (input$plot_type %in% c("bar", "clustered_bar")) {
        y_choices <- c("None" = "none", names(df))
      }
      updateSelectInput(session, "y_var", choices = y_choices)
      
      updateSelectInput(session, "color_var", choices = c("None" = "none", names(df)))
      updateSelectInput(session, "size_var", choices = c("None" = "none", names(df)))
      updateSelectInput(session, "shape_var", choices = c("None" = "none", names(df)))
      updateSelectInput(session, "facet_var", choices = c("None" = "none", names(df)))
      updateSelectInput(session, "group_var", choices = c("None" = "none", names(df)))
      updateSelectInput(session, "pie_var", choices = names(df))
      updateSelectInput(session, "waffle_var", choices = names(df))
      updateSelectInput(session, "word_var", choices = names(df))
      updateSelectInput(session, "freq_var", choices = names(df))
      updateSelectInput(session, "id_var", choices = names(df))
      updateSelectInput(session, "cluster_var", choices = names(df))
    }
  })
  
  # Generate plot
  plot_obj <- eventReactive(input$run_analysis, {
    req(data())
    df <- data()
    if (is.null(df)) return(NULL)
    
    # Base plot
    p <- ggplot(data = df)
    
    # Get colors based on selection
    if (input$color_choice == "single") {
      color <- input$single_color
    } else {
      color <- category_colors()
    }
    
    # Add geom based on plot type
    if (input$plot_type == "scatter") {
      if (input$color_choice == "single") {
        p <- p + geom_point(
          aes(
            x = .data[[input$x_var]],
            y = .data[[input$y_var]],
            size = if (input$size_var != "none") .data[[input$size_var]] else NULL,
            shape = if (input$shape_var != "none") .data[[input$shape_var]] else NULL,
            group = if (input$group_var != "none") .data[[input$group_var]] else NULL
          ),
          alpha = input$point_alpha,
          size = input$point_size,
          color = color
        )
      } else {
        p <- p + geom_point(
          aes(
            x = .data[[input$x_var]],
            y = .data[[input$y_var]],
            color = if (input$color_var != "none") .data[[input$color_var]] else NULL,
            size = if (input$size_var != "none") .data[[input$size_var]] else NULL,
            shape = if (input$shape_var != "none") .data[[input$shape_var]] else NULL,
            group = if (input$group_var != "none") .data[[input$group_var]] else NULL
          ),
          alpha = input$point_alpha,
          size = input$point_size
        )
      }
      
      # Add correlation if requested
      if (input$show_correlation) {
        p <- p + ggpubr::stat_cor(
          method = input$cor_method,
          label.x.npc = "left",
          label.y.npc = "top"
        )
      }
      
      # Add data point IDs if requested
      if (input$show_ids && input$id_var != "none") {
        p <- p + geom_text_repel(
          aes(
            x = .data[[input$x_var]],
            y = .data[[input$y_var]],
            label = .data[[input$id_var]]
          ),
          size = 3,
          max.overlaps = 20
        )
      }
    } else if (input$plot_type == "line") {
      if (input$color_choice == "single") {
        p <- p + geom_line(
          aes(
            x = .data[[input$x_var]],
            y = .data[[input$y_var]],
            size = if (input$size_var != "none") .data[[input$size_var]] else NULL,
            linetype = if (input$shape_var != "none") .data[[input$shape_var]] else NULL,
            group = if (input$group_var != "none") .data[[input$group_var]] else NULL
          ),
          size = input$line_size,
          alpha = input$line_alpha,
          linetype = input$line_type,
          color = color
        )
      } else {
        p <- p + geom_line(
          aes(
            x = .data[[input$x_var]],
            y = .data[[input$y_var]],
            color = if (input$color_var != "none") .data[[input$color_var]] else NULL,
            size = if (input$size_var != "none") .data[[input$size_var]] else NULL,
            linetype = if (input$shape_var != "none") .data[[input$shape_var]] else NULL,
            group = if (input$group_var != "none") .data[[input$group_var]] else NULL
          ),
          size = input$line_size,
          alpha = input$line_alpha,
          linetype = input$line_type
        )
      }
    } else if (input$plot_type == "bar") {
      # Prepare data for bar chart
      if (input$y_var == "none") {
        # Count plot - use the improved approach from the example
        bar_data <- df %>%
          group_by(.data[[input$x_var]]) %>%
          summarise(Frequency = n()) %>%
          mutate(Percent = round(100 * Frequency / sum(Frequency), input$value_digits))
        
        # Combine frequency and percent for a compact label
        if (input$show_percentage) {
          if (input$show_both) {
            bar_data$Label <- paste0(bar_data$Frequency, "\n", bar_data$Percent, "%")
          } else {
            bar_data$Label <- paste0(bar_data$Percent, "%")
          }
        } else {
          bar_data$Label <- bar_data$Frequency
        }
        
        if (input$color_choice == "single") {
          p <- ggplot(bar_data, aes(x = .data[[input$x_var]], y = Frequency)) +
            geom_bar(
              stat = "identity",
              position = input$position,
              alpha = input$fill_alpha,
              fill = color,
              color = "black"  # Add black border like in the example
            )
          
          # Add values on top of bars if requested
          if (input$show_values) {
            p <- p + geom_text(
              aes(label = Label),
              vjust = -0.3,
              size = input$value_size
            ) +
              ylim(0, max(bar_data$Frequency) * 1.15)  # Give space above bars for labels
          }
        } else {
          # Determine which variable to use for fill
          fill_var <- if (input$color_var != "none") input$color_var else input$x_var
          
          p <- ggplot(bar_data, aes(x = .data[[input$x_var]], y = Frequency, fill = .data[[fill_var]])) +
            geom_bar(
              stat = "identity",
              position = input$position,
              alpha = input$fill_alpha,
              color = "black"  # Add black border like in the example
            )
          
          # Add values on top of bars if requested
          if (input$show_values) {
            p <- p + geom_text(
              aes(label = Label),
              position = position_dodge(width = 0.9),
              vjust = -0.3,
              size = input$value_size
            ) +
              ylim(0, max(bar_data$Frequency) * 1.15)  # Give space above bars for labels
          }
        }
      } else {
        # Bar plot with y variable - keep existing functionality
        if (input$color_choice == "single") {
          p <- p + geom_bar(
            aes(
              x = .data[[input$x_var]],
              y = .data[[input$y_var]],
              group = if (input$group_var != "none") .data[[input$group_var]] else NULL
            ),
            stat = "identity",
            position = input$position,
            alpha = input$fill_alpha,
            fill = color,
            color = "black"  # Add black border like in the example
          )
          
          # Add values on top of bars if requested
          if (input$show_values) {
            p <- p + geom_text(
              aes(
                x = .data[[input$x_var]],
                y = .data[[input$y_var]],
                label = round(.data[[input$y_var]], input$value_digits)
              ),
              vjust = -0.3,
              size = input$value_size,
              position = position_dodge(width = 0.9)
            )
          }
        } else {
          p <- p + geom_bar(
            aes(
              x = .data[[input$x_var]],
              y = .data[[input$y_var]],
              fill = if (input$color_var != "none") .data[[input$color_var]] else NULL,
              group = if (input$group_var != "none") .data[[input$group_var]] else NULL
            ),
            stat = "identity",
            position = input$position,
            alpha = input$fill_alpha,
            color = "black"  # Add black border like in the example
          )
          
          # Add values on top of bars if requested
          if (input$show_values) {
            p <- p + geom_text(
              aes(
                x = .data[[input$x_var]],
                y = .data[[input$y_var]],
                label = round(.data[[input$y_var]], input$value_digits),
                group = if (input$color_var != "none") .data[[input$color_var]] else NULL
              ),
              vjust = -0.3,
              size = input$value_size,
              position = position_dodge(width = 0.9)
            )
          }
        }
      }
      
      
      # Rotate bar plot if requested
      if (input$rotate_barplot) {
        p <- p + coord_flip()
      }
    } else if (input$plot_type == "clustered_bar") {
      req(input$x_var, input$cluster_var)
      
      # Prepare data for clustered bar chart
      if (input$y_var == "none") {
        # Count plot
        bar_data <- df %>%
          group_by(.data[[input$x_var]], .data[[input$cluster_var]]) %>%
          summarise(count = n(), .groups = "drop") %>%
          group_by(.data[[input$x_var]]) %>%
          mutate(percentage = count / sum(count) * 100)
        
        # Create labels based on user selection
        if (input$show_values) {
          if (input$show_percentage) {
            if (input$show_both) {
              bar_data <- bar_data %>%
                mutate(label = paste0(count, "\n(", sprintf(paste0("%.", input$value_digits, "f"), percentage), "%)"))
            } else {
              bar_data <- bar_data %>%
                mutate(label = sprintf(paste0("%.", input$value_digits, "f"), percentage))
            }
          } else {
            bar_data <- bar_data %>%
              mutate(label = count)
          }
        }
        
        if (input$color_choice == "single") {
          p <- ggplot(bar_data, aes(x = .data[[input$x_var]], y = count, fill = .data[[input$cluster_var]])) +
            geom_bar(
              stat = "identity",
              position = input$position,
              alpha = input$fill_alpha
            ) +
            scale_fill_manual(values = rep(color, length(unique(bar_data[[input$cluster_var]]))))
        } else {
          p <- ggplot(bar_data, aes(x = .data[[input$x_var]], y = count, fill = .data[[input$cluster_var]])) +
            geom_bar(
              stat = "identity",
              position = input$position,
              alpha = input$fill_alpha
            )
        }
        
        # Add values on top of bars if requested
        if (input$show_values) {
          p <- p + geom_text(
            aes(label = label),
            position = position_dodge(width = 0.9),
            vjust = -0.5,
            size = input$value_size
          )
        }
      } else {
        # Bar plot with y variable
        bar_data <- df %>%
          group_by(.data[[input$x_var]], .data[[input$cluster_var]]) %>%
          summarise(y_value = sum(.data[[input$y_var]]), .groups = "drop") %>%
          group_by(.data[[input$x_var]]) %>%
          mutate(percentage = y_value / sum(y_value) * 100)
        
        # Create labels based on user selection
        if (input$show_values) {
          if (input$show_percentage) {
            if (input$show_both) {
              bar_data <- bar_data %>%
                mutate(label = paste0(round(y_value, input$value_digits), "\n(", sprintf(paste0("%.", input$value_digits, "f"), percentage), "%)"))
            } else {
              bar_data <- bar_data %>%
                mutate(label = sprintf(paste0("%.", input$value_digits, "f"), percentage))
            }
          } else {
            bar_data <- bar_data %>%
              mutate(label = round(y_value, input$value_digits))
          }
        }
        
        if (input$color_choice == "single") {
          p <- ggplot(bar_data, aes(x = .data[[input$x_var]], y = y_value, fill = .data[[input$cluster_var]])) +
            geom_bar(
              stat = "identity",
              position = input$position,
              alpha = input$fill_alpha
            ) +
            scale_fill_manual(values = rep(color, length(unique(bar_data[[input$cluster_var]]))))
        } else {
          p <- ggplot(bar_data, aes(x = .data[[input$x_var]], y = y_value, fill = .data[[input$cluster_var]])) +
            geom_bar(
              stat = "identity",
              position = input$position,
              alpha = input$fill_alpha
            )
        }
        
        # Add values on top of bars if requested
        if (input$show_values) {
          p <- p + geom_text(
            aes(label = label),
            position = position_dodge(width = 0.9),
            vjust = -0.5,
            size = input$value_size
          )
        }
      }
      
      # Rotate bar plot if requested
      if (input$rotate_barplot) {
        p <- p + coord_flip()
      }
    } else if (input$plot_type == "histogram") {
      if (input$color_choice == "single") {
        p <- p + geom_histogram(
          aes(
            x = .data[[input$x_var]],
            group = if (input$group_var != "none") .data[[input$group_var]] else NULL
          ),
          binwidth = input$binwidth,
          bins = input$bins,
          fill = color,
          alpha = input$fill_alpha,
          position = input$position
        )
      } else {
        p <- p + geom_histogram(
          aes(
            x = .data[[input$x_var]],
            fill = if (input$color_var != "none") .data[[input$color_var]] else NULL,
            group = if (input$group_var != "none") .data[[input$group_var]] else NULL
          ),
          binwidth = input$binwidth,
          bins = input$bins,
          alpha = input$fill_alpha,
          position = input$position
        )
      }
    } else if (input$plot_type == "density") {
      if (input$color_choice == "single") {
        p <- p + geom_density(
          aes(
            x = .data[[input$x_var]],
            group = if (input$group_var != "none") .data[[input$group_var]] else NULL
          ),
          fill = color,
          alpha = input$fill_alpha,
          adjust = input$binwidth
        )
      } else {
        p <- p + geom_density(
          aes(
            x = .data[[input$x_var]],
            fill = if (input$color_var != "none") .data[[input$color_var]] else NULL,
            group = if (input$group_var != "none") .data[[input$group_var]] else NULL
          ),
          alpha = input$fill_alpha,
          adjust = input$binwidth
        )
      }
    } else if (input$plot_type == "boxplot") {
      if (input$color_choice == "single") {
        p <- p + geom_boxplot(
          aes(
            x = if (is.numeric(df[[input$x_var]])) NULL else .data[[input$x_var]],
            y = .data[[input$y_var]],
            group = if (input$group_var != "none") .data[[input$group_var]] else NULL
          ),
          fill = color,
          alpha = input$fill_alpha,
          notch = input$notch,
          width = input$width,
          outlier.shape = if (input$show_outliers) 19 else NA
        )
      } else {
        p <- p + geom_boxplot(
          aes(
            x = if (is.numeric(df[[input$x_var]])) NULL else .data[[input$x_var]],
            y = .data[[input$y_var]],
            fill = if (input$color_var != "none") .data[[input$color_var]] else NULL,
            group = if (input$group_var != "none") .data[[input$group_var]] else NULL
          ),
          alpha = input$fill_alpha,
          notch = input$notch,
          width = input$width,
          outlier.shape = if (input$show_outliers) 19 else NA
        )
      }
      
      # Show outliers with labels if requested
      if (input$show_outliers && input$id_var != "none") {
        outliers <- df %>%
          group_by(across(c(input$x_var, input$color_var))) %>%
          mutate(
            iqr = IQR(.data[[input$y_var]], na.rm = TRUE),
            q1 = quantile(.data[[input$y_var]], 0.25, na.rm = TRUE),
            q3 = quantile(.data[[input$y_var]], 0.75, na.rm = TRUE),
            is_outlier = .data[[input$y_var]] < (q1 - 1.5 * iqr) | .data[[input$y_var]] > (q3 + 1.5 * iqr)
          ) %>%
          filter(is_outlier)
        
        if (nrow(outliers) > 0) {
          p <- p + geom_text_repel(
            data = outliers,
            aes(
              x = if (is.numeric(df[[input$x_var]])) 1 else .data[[input$x_var]],
              y = .data[[input$y_var]],
              label = .data[[input$id_var]]
            ),
            size = 3,
            max.overlaps = 20
          )
        }
      }
      
      # Rotate plot if requested
      if (input$rotate_plot) {
        p <- p + coord_flip()
      }
    } else if (input$plot_type == "violin") {
      if (input$color_choice == "single") {
        p <- p + geom_violin(
          aes(
            x = if (is.numeric(df[[input$x_var]])) NULL else .data[[input$x_var]],
            y = .data[[input$y_var]],
            group = if (input$group_var != "none") .data[[input$group_var]] else NULL
          ),
          fill = color,
          alpha = input$fill_alpha,
          trim = input$trim,
          scale = input$violin_scale,
          width = input$width
        )
      } else {
        p <- p + geom_violin(
          aes(
            x = if (is.numeric(df[[input$x_var]])) NULL else .data[[input$x_var]],
            y = .data[[input$y_var]],
            fill = if (input$color_var != "none") .data[[input$color_var]] else NULL,
            group = if (input$group_var != "none") .data[[input$group_var]] else NULL
          ),
          alpha = input$fill_alpha,
          trim = input$trim,
          scale = input$violin_scale,
          width = input$width
        )
      }
      
      # Rotate plot if requested
      if (input$rotate_plot) {
        p <- p + coord_flip()
      }
    } else if (input$plot_type == "pie") {
      pie_data <- df %>%
        group_by(.data[[input$pie_var]]) %>%
        summarise(count = n()) %>%
        mutate(percentage = count / sum(count) * 100,
               ypos = cumsum(percentage) - 0.5 * percentage)
      
      if (input$color_choice == "single") {
        p <- ggplot(pie_data, aes(x = "", y = percentage, fill = .data[[input$pie_var]])) +
          geom_bar(stat = "identity", width = input$pie_size, alpha = input$fill_alpha) +
          coord_polar("y", start = 0) +
          scale_fill_manual(values = rep(color, nrow(pie_data)))
      } else {
        p <- ggplot(pie_data, aes(x = "", y = percentage, fill = .data[[input$pie_var]])) +
          geom_bar(stat = "identity", width = input$pie_size, alpha = input$fill_alpha) +
          coord_polar("y", start = 0)
      }
      
      if (input$pie_labels) {
        p <- p + geom_text(aes(y = ypos, label = paste0(round(percentage, input$label_digits), "%")),
                           color = "white", size = input$label_size)
      }
    } else if (input$plot_type == "donut") {
      pie_data <- df %>%
        group_by(.data[[input$pie_var]]) %>%
        summarise(count = n()) %>%
        mutate(percentage = count / sum(count) * 100,
               ypos = cumsum(percentage) - 0.5 * percentage)
      
      if (input$color_choice == "single") {
        p <- ggplot(pie_data, aes(x = 2, y = percentage, fill = .data[[input$pie_var]])) +
          geom_bar(stat = "identity", width = input$pie_size, alpha = input$fill_alpha) +
          coord_polar("y", start = 0) +
          xlim(0.5, 2.5) +
          scale_fill_manual(values = rep(color, nrow(pie_data)))
      } else {
        p <- ggplot(pie_data, aes(x = 2, y = percentage, fill = .data[[input$pie_var]])) +
          geom_bar(stat = "identity", width = input$pie_size, alpha = input$fill_alpha) +
          coord_polar("y", start = 0) +
          xlim(0.5, 2.5)
      }
      
      if (input$pie_labels) {
        p <- p + geom_text(aes(y = ypos, label = paste0(round(percentage, input$label_digits), "%")),
                           color = "white", size = input$label_size)
      }
    } else if (input$plot_type == "waffle") {
      waffle_data <- df %>%
        group_by(.data[[input$waffle_var]]) %>%
        summarise(count = n()) %>%
        mutate(percent = round(count / sum(count) * 100))
      
      if (input$color_choice == "single") {
        p <- ggplot(waffle_data, aes(fill = .data[[input$waffle_var]], values = percent)) +
          geom_waffle(n_rows = input$waffle_rows, size = 0.33, color = "white", flip = TRUE) +
          scale_fill_manual(values = rep(color, nrow(waffle_data))) +
          coord_equal() +
          theme_void()
      } else {
        p <- ggplot(waffle_data, aes(fill = .data[[input$waffle_var]], values = percent)) +
          geom_waffle(n_rows = input$waffle_rows, size = 0.33, color = "white", flip = TRUE) +
          coord_equal() +
          theme_void()
      }
    } else if (input$plot_type == "wordcloud") {
      word_data <- df %>%
        group_by(.data[[input$word_var]]) %>%
        summarise(freq = sum(.data[[input$freq_var]]))
      
      p <- ggplot(word_data, aes(label = .data[[input$word_var]], size = freq)) +
        geom_text_wordcloud() +
        scale_size_area(max_size = 20) +
        theme_minimal()
    }
    
    # Apply color scale if custom colors are provided
    if (input$color_choice == "category" && !is.null(color)) {
      p <- p + scale_fill_manual(values = color)
      p <- p + scale_color_manual(values = color)
    }
    
    # Add facet if requested
    if (input$facet_var != "none") {
      p <- p + facet_wrap(as.formula(paste("~", input$facet_var)))
    }
    
    # Add labels and titles
    p <- p + labs(
      title = input$title,
      x = if (input$x_lab != "") input$x_lab else input$x_var,
      y = if (input$y_lab != "") input$y_lab else input$y_var,
      caption = input$caption,
      tag = input$tag
    )
    
    # Apply theme
    theme_func <- switch(input$theme,
                         "theme_gray" = theme_gray,
                         "theme_classic" = theme_classic,
                         "theme_minimal" = theme_minimal,
                         "theme_bw" = theme_bw,
                         "theme_light" = theme_light,
                         "theme_dark" = theme_dark,
                         "theme_void" = theme_void,
                         "theme_test" = theme_test,
                         "theme_pubr" = ggpubr::theme_pubr,
                         "theme_pubclean" = ggpubr::theme_pubclean,
                         "theme_linedraw" = theme_linedraw)
    
    p <- p + theme_func(base_size = input$base_size)
    
    # Flip coordinates if requested
    if (input$coord_flip) {
      p <- p + coord_flip()
    }
    
    return(p)
  })
  
  # Statistical test results
  output$statistical_results <- renderPrint({
    req(input$run_tests > 0)  # Only run when button is clicked
    df <- data()
    if (is.null(df)) return(NULL)
    
    cat("=== Statistical Test Results ===\n\n")
    
    # Correlation test for scatter plots
    if (input$plot_type == "scatter" && input$show_correlation) {
      cat("Correlation Test:\n")
      cor_test <- cor.test(df[[input$x_var]], df[[input$y_var]], 
                           method = input$cor_method)
      print(cor_test)
      cat("\n")
    }
    
    # Statistical test for boxplots and violin plots
    if ((input$plot_type == "boxplot" || input$plot_type == "violin") && input$show_test) {
      if (input$color_var != "none") {
        formula <- as.formula(paste(input$y_var, "~", input$color_var))
      } else {
        formula <- as.formula(paste(input$y_var, "~", input$x_var))
      }
      
      cat("Statistical Test:\n")
      test_result <- switch(input$test_method,
                            "wilcox.test" = wilcox.test(formula, data = df),
                            "t.test" = t.test(formula, data = df),
                            "kruskal.test" = kruskal.test(formula, data = df))
      print(test_result)
      cat("\n")
    }
    
    # Chi-square test for bar plots
    if (input$plot_type == "bar" && input$show_chisq && input$x_var != "none" && input$y_var != "none") {
      cat("Chi-square Test:\n")
      chi_test <- chisq.test(df[[input$x_var]], df[[input$y_var]])
      print(chi_test)
      cat("\n")
    }
    
    # Normality tests for histograms
    if (input$plot_type == "histogram" && input$show_normality) {
      cat("Normality Tests:\n")
      cat("Anderson-Darling Test:\n")
      norm_test <- nortest::ad.test(df[[input$x_var]])
      print(norm_test)
      
      cat("\nShapiro-Wilk Test (for n < 5000):\n")
      if (length(df[[input$x_var]]) < 5000) {
        print(shapiro.test(df[[input$x_var]]))
      } else {
        cat("Sample size too large for Shapiro-Wilk test (n >= 5000)\n")
      }
    }
  })
  
  # Render static plot
  output$static_plot <- renderPlot({
    p <- plot_obj()
    if (is.null(p)) return(NULL)
    print(p)
  })
  
  # Render interactive plot
  output$interactive_plot <- renderGirafe({
    p <- plot_obj()
    if (is.null(p)) return(NULL)
    
    # Convert to interactive plot
    girafe(
      ggobj = p,
      width_svg = input$plot_width,
      height_svg = input$plot_height,
      options = list(
        opts_tooltip(css = "background-color:white;color:black;padding:5px;border-radius:5px;"),
        opts_hover(css = "stroke-width:2px;"),
        opts_selection(type = "none")
      )
    )
  })
  
  # Display data table
  output$data_table <- renderDT({
    df <- data()
    if (is.null(df)) return(NULL)
    datatable(
      df,
      extensions = c("Buttons", "Scroller"),
      options = list(
        dom = "Bfrtip",
        buttons = c("copy", "csv", "excel", "pdf", "print"),
        scrollX = TRUE,
        scrollY = "500px",
        scroller = TRUE
      )
    )
  })
  
  # Generate plot code
  output$plot_code <- renderText({
    p <- plot_obj()
    if (is.null(p)) return("No plot generated yet.")
    
    # Capture the plot code
    plot_code <- capture.output(print(p))
    
    # Combine into a single string
    paste(plot_code, collapse = "\n")
  })
  
  # Download plot
  output$download_plot <- downloadHandler(
    filename = function() {
      paste("plot-", Sys.Date(), ".png", sep = "")
    },
    content = function(file) {
      p <- plot_obj()
      if (is.null(p)) return(NULL)
      
      ggsave(
        file,
        plot = p,
        width = input$plot_width,
        height = input$plot_height,
        dpi = input$plot_dpi,
        device = "png"
      )
    }
  )
  
  # Download data
  output$download_data <- downloadHandler(
    filename = function() {
      paste("data-", Sys.Date(), ".csv", sep = "")
    },
    content = function(file) {
      df <- data()
      if (is.null(df)) return(NULL)
      write.csv(df, file, row.names = FALSE)
    }
  )
}

# Run the application
shinyApp(ui = ui, server = server)
