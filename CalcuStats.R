library(shiny)
library(shinythemes)
library(officer)
library(flextable)
library(magrittr)
library(pwr)
library(e1071)

counter_file <- "counter.rds"
if (!file.exists(counter_file)) saveRDS(0, counter_file)
counter <- readRDS(counter_file) + 1
saveRDS(counter, counter_file)

getGreeting <- function() {
  hour <- as.numeric(format(Sys.time(), "%H"))
  if (hour < 12) {
    return("Good morning!")
  } else if (hour < 18) {
    return("Good afternoon!")
  } else {
    return("Good evening!")
  }
}

ui <- navbarPage(
  title = div(icon("chart-bar"), "CalcuStats"),
  theme = shinytheme("cosmo"),
  footer = div(
    style = "text-align: center; padding: 10px; background-color: #f8f9fa; border-top: 1px solid #ddd;",
    HTML("<p>For suggestions or support, contact <a href='mailto:mudassiribrahim30@gmail.com'>mudassiribrahim30@gmail.com</a></p>"),
    div(
      style = "margin-top: 10px; font-size: 0.9em; color: #555;",
      "Developed by Mudasir Mohammed Ibrahim (Registered Nurse, BSc, Dip)"
    )
  ),
  
  tabPanel("Taro Yamane",
           sidebarLayout(
             sidebarPanel(
               width = 4,
               numericInput("e", "Margin of Error (e)", value = 0.05, min = 0.0001, max = 1, step = 0.01),
               numericInput("non_response", "Non-response Rate (%)", value = 10, min = 0, max = 100, step = 1),
               actionButton("addStratum", "Add Stratum", class = "btn-primary"),
               actionButton("removeStratum", "Remove Last Stratum", class = "btn-warning"),
               uiOutput("stratumInputs"),
               div(
                 style = "margin-top: 10px; padding: 10px; background-color: #f8f9fa; border-radius: 5px;",
                 p("For a single population, enter the total size below."),
                 p("For multiple populations, click 'Add Stratum' to create additional groups.")
               ),
               downloadButton("downloadWord", "Download as Word", class = "btn-success")
             ),
             mainPanel(
               width = 8,
               h3("Taro Yamane Method with Proportional Allocation"),
               h4(textOutput("greetingText"), style = "color: #1A5276; font-weight: bold;"),
               
               div(
                 style = "padding: 10px; background-color: #f0f8ff; border-left: 5px solid #1A5276; margin-bottom: 20px;",
                 HTML("
                 <strong>Taro Yamane Formula:</strong><br>
                 <em>n = N / (1 + N &times; e²)</em><br>
                 <ul>
                   <li><strong>n</strong> = required sample size</li>
                   <li><strong>N</strong> = population size</li>
                   <li><strong>e</strong> = margin of error (in decimal)</li>
                 </ul>
                 This formula assumes a 95% confidence level and simplifies sample size estimation for large populations.
               ")
               ),
               
               fluidRow(
                 column(6, verbatimTextOutput("totalPop")),
                 column(6, verbatimTextOutput("sampleSize"))
               ),
               verbatimTextOutput("formulaExplanation"),
               verbatimTextOutput("adjustedSampleSize"),
               h4("Proportional Allocation"),
               p("The proportional allocation formula is:"),
               p(HTML("<em>n<sub>i</sub> = (N<sub>i</sub> / N) &times; n</em>")),
               tableOutput("allocationTable"),
               textOutput("interpretationText")
             )
           )
  ),
  
  tabPanel("Cochran Formula",
           sidebarLayout(
             sidebarPanel(
               width = 4,
               numericInput("p", "Estimated Proportion (p)", value = 0.5, min = 0.01, max = 0.99, step = 0.01),
               numericInput("z", "Z-score (Z)", value = 1.96),
               numericInput("e_c", "Margin of Error (e)", value = 0.05, min = 0.0001, max = 1, step = 0.01),
               numericInput("N_c", "Population Size (optional)", value = NULL, min = 1)
             ),
             mainPanel(
               width = 8,
               h3("Cochran Sample Size Calculator"),
               div(
                 style = "padding: 10px; background-color: #f0f8ff; border-left: 5px solid #1A5276; margin-bottom: 20px;",
                 HTML("
                 <strong>Cochran's Formula:</strong><br>
                 <em>n₀ = (Z² × p × (1-p)) / e²</em><br>
                 <ul>
                   <li><strong>n₀</strong> = required sample size</li>
                   <li><strong>Z</strong> = Z-score for desired confidence level</li>
                   <li><strong>p</strong> = estimated proportion of the attribute present in the population</li>
                   <li><strong>e</strong> = margin of error (in decimal)</li>
                 </ul>
                 For finite populations, apply the finite population correction: n = n₀ / (1 + (n₀ - 1)/N)
               ")
               ),
               verbatimTextOutput("cochranFormulaExplanation"),
               verbatimTextOutput("cochranSample"),
               textOutput("cochranInterpretation")
             )
           )
  ),
  
  tabPanel("Power Analysis",
           sidebarLayout(
             sidebarPanel(
               width = 4,
               selectInput("testType", "Statistical Test:",
                          choices = c("Independent t-test", "Paired t-test", "One-sample t-test", "One-Way ANOVA",
                                     "Two-Way ANOVA", "Proportion", "Correlation", "Chi-squared",
                                     "Simple Linear Regression", "Multiple Linear Regression")),
               numericInput("effectSize", "Effect Size", value = 0.5),
               numericInput("alpha", "Significance Level (alpha)", value = 0.05),
               numericInput("power", "Desired Power", value = 0.8),
               conditionalPanel(
                 condition = "input.testType == 'Multiple Linear Regression'",
                 numericInput("predictors", "Number of Predictors", value = 2, min = 1)
               ),
               actionButton("runPower", "Run Power Analysis", class = "btn-primary")
             ),
             mainPanel(
               width = 8,
               h3("Power Analysis for Inferential Tests"),
               verbatimTextOutput("powerResult")
             )
           )
  ),
  
  tabPanel("Descriptive Statistics",
           sidebarLayout(
             sidebarPanel(
               width = 4,
               tags$textarea(id = "dataInput", rows = 10, cols = 30,
                            placeholder = "Paste a column of data (with header) from Excel or statistical software..."),
               actionButton("runDesc", "Get Descriptive Statistics", class = "btn-primary")
             ),
             mainPanel(
               width = 8,
               h3("Descriptive Statistics"),
               verbatimTextOutput("descResult")
             )
           )
  ),
  
  tabPanel("Usage",
           fluidPage(
             div(
               style = "text-align: center; padding: 50px;",
               h3("App Usage Statistics"),
               div(
                 style = "font-size: 24px; margin: 20px; padding: 20px; background-color: #f8f9fa; border-radius: 10px;",
                 paste("This app has been used", counter, "times.")
               ),
               p("Thank you for using CalcuStats!")
             )
           )
  )
)

server <- function(input, output, session) {
  output$greetingText <- renderText({
    paste(getGreeting(), "Welcome to CalcuStats.")
  })
  
  rv <- reactiveValues(stratumCount = 1)
  
  observeEvent(input$addStratum, { rv$stratumCount <- rv$stratumCount + 1 })
  observeEvent(input$removeStratum, { if (rv$stratumCount > 1) rv$stratumCount <- rv$stratumCount - 1 })
  
  output$stratumInputs <- renderUI({
    lapply(1:rv$stratumCount, function(i) {
      tagList(
        div(style = "margin-bottom: 10px; width: 100%;",
            textInput(paste0("stratum", i),
                      label = paste("Stratum", i, "Name"),
                      value = paste("Stratum", i),
                      width = "100%")
        ),
        div(style = "margin-bottom: 20px; width = 100%;",
            numericInput(paste0("pop", i),
                         label = paste("Stratum", i, "Population"),
                         value = 100,
                         min = 0,
                         width = "100%")
        ),
        tags$hr()
      )
    })
  })
  
  totalPopulation <- reactive({
    sum(sapply(1:rv$stratumCount, function(i) input[[paste0("pop", i)]]), na.rm = TRUE)
  })
  
  sampleSize <- reactive({
    N <- totalPopulation()
    e <- input$e
    if (N == 0 || e == 0) return(0)
    ceiling(N / (1 + N * e^2))
  })
  
  adjustedSampleSize <- reactive({
    n <- sampleSize()
    if (n == 0) return(0)
    non_response_rate <- input$non_response / 100
    ceiling(n / (1 - non_response_rate))
  })
  
  output$formulaExplanation <- renderText({
    N <- totalPopulation()
    e <- input$e
    if (N == 0 || e == 0) return("")
    
    denominator <- 1 + N * e^2
    raw_sample <- N / denominator
    
    paste(
      "Yamane Formula Calculation Steps:\n",
      "1. Formula: n = N / (1 + N*e²)\n",
      "2. Total Population (N) = ", N, "\n",
      "3. Margin of Error (e) = ", e, "\n",
      "4. Calculate denominator: 1 + N*e² = 1 + (", N, " * ", e, "²) = 1 + (", N, " * ", e^2, ") = ", denominator, "\n",
      "5. Calculate raw sample size: ", N, " / ", denominator, " = ", raw_sample, "\n",
      "6. Apply ceiling function to round up to nearest integer: ", ceiling(raw_sample), "\n",
      "\nFinal calculated sample size after rounding: ", sampleSize()
    )
  })
  
  output$adjustedSampleSize <- renderText({
    n <- sampleSize()
    adj_n <- adjustedSampleSize()
    non_response_rate <- input$non_response
    
    paste(
      "Adjusting for Non-Response:\n",
      "1. Original sample size: ", n, "\n",
      "2. Non-response rate: ", non_response_rate, "%\n",
      "3. Adjusted sample size formula: n_adjusted = n / (1 - non_response_rate)\n",
      "4. Calculation: ", n, " / (1 - ", non_response_rate/100, ") = ", n / (1 - non_response_rate/100), "\n",
      "5. Apply ceiling function: ", adj_n, "\n",
      "\nFinal adjusted sample size: ", adj_n
    )
  })
  
  allocationData <- reactive({
    N <- totalPopulation()
    n <- adjustedSampleSize()
    strata <- sapply(1:rv$stratumCount, function(i) input[[paste0("stratum", i)]] )
    pops <- sapply(1:rv$stratumCount, function(i) input[[paste0("pop", i)]] )
    proportions <- pops / N
    raw_samples <- proportions * n
    rounded_samples <- floor(raw_samples)
    remainder <- n - sum(rounded_samples)
    
    if (remainder > 0) {
      decimal_parts <- raw_samples - rounded_samples
      indices <- order(decimal_parts, decreasing = TRUE)[1:remainder]
      rounded_samples[indices] <- rounded_samples[indices] + 1
    }
    
    calculations <- sapply(1:rv$stratumCount, function(i) {
      paste0(
        "(", pops[i], " / ", N, ") × ", n, 
        " = ", round(proportions[i], 4), " × ", n, 
        " = ", round(raw_samples[i], 2), " → ", rounded_samples[i]
      )
    })
    
    df <- data.frame(
      Stratum = strata,
      Population = pops,
      Calculation = calculations,
      Proportional_Sample = rounded_samples
    )
    
    total_row <- data.frame(
      Stratum = "Total",
      Population = N,
      Calculation = paste0("Sum = ", sum(rounded_samples)),
      Proportional_Sample = sum(rounded_samples)
    )
    
    rbind(df, total_row)
  })
  
  interpretationText <- reactive({
    alloc <- allocationData()
    if (is.null(alloc)) return("")
    sentences <- paste0(alloc$Stratum[-nrow(alloc)],
                        " (population: ", alloc$Population[-nrow(alloc)],
                        ") should contribute ", alloc$Proportional_Sample[-nrow(alloc)],
                        " participants.")
    paste0("Based on a total population of ", totalPopulation(),
           ", a margin of error of ", input$e,
           ", and a non-response rate of ", input$non_response, "%",
           ", the required sample size is ", adjustedSampleSize(), ". ",
           paste(sentences, collapse = " "))
  })
  
  output$totalPop <- renderText({ paste("Total Population:", totalPopulation()) })
  output$sampleSize <- renderText({ paste("Initial Sample Size:", sampleSize()) })
  output$allocationTable <- renderTable({ allocationData() })
  output$interpretationText <- renderText({ interpretationText() })
  
  output$downloadWord <- downloadHandler(
    filename = function() paste("Sample_Size_Report_", Sys.Date(), ".docx", sep = ""),
    content = function(file) {
      doc <- read_docx() %>%
        body_add_par("StatMystery unlocked with Mudasir", style = "heading 1") %>%
        body_add_par("Sample Size Calculation (Taro Yamane Method)", style = "heading 2") %>%
        body_add_par(paste("Total Population:", totalPopulation()), style = "Normal") %>%
        body_add_par(paste("Margin of Error:", input$e), style = "Normal") %>%
        body_add_par(paste("Non-response Rate:", input$non_response, "%"), style = "Normal") %>%
        body_add_par(paste("Initial Sample Size:", sampleSize()), style = "Normal") %>%
        body_add_par(paste("Adjusted Sample Size (accounting for non-response):", adjustedSampleSize()), style = "Normal") %>%
        body_add_par("Proportional Allocation Table:", style = "heading 2") %>%
        body_add_par("The proportional allocation formula is: n_i = (N_i / N) × n", style = "Normal")
      ft <- flextable(allocationData()) %>% autofit()
      doc <- doc %>% body_add_flextable(ft)
      doc <- doc %>%
        body_add_par("Interpretation", style = "heading 2") %>%
        body_add_par(interpretationText(), style = "Normal")
      print(doc, target = file)
    }
  )
  
  cochranSampleSize <- reactive({
    p <- input$p
    z <- input$z
    e <- input$e_c
    N <- input$N_c
    
    n0 <- (z^2 * p * (1 - p)) / (e^2)
    
    if (!is.null(N) && !is.na(N) && N > 0) {
      n <- n0 / (1 + (n0 - 1)/N)
      return(list(initial = n0, corrected = n, finite_correction_applied = TRUE))
    } else {
      return(list(initial = n0, corrected = n0, finite_correction_applied = FALSE))
    }
  })
  
  output$cochranFormulaExplanation <- renderText({
    p <- input$p
    z <- input$z
    e <- input$e_c
    N <- input$N_c
    res <- cochranSampleSize()
    
    calculation_steps <- paste(
      "Cochran's Formula Calculation Steps:\n",
      "1. Basic formula: n₀ = (Z² × p × (1-p)) / e²\n",
      "2. Z-score (Z) = ", z, "\n",
      "3. Estimated proportion (p) = ", p, "\n",
      "4. Margin of error (e) = ", e, "\n",
      "5. Calculate numerator (Z² × p × (1-p)): (", z, "² × ", p, " × ", (1-p), ") = ", 
      (z^2 * p * (1 - p)), "\n",
      "6. Calculate denominator (e²): ", e, "² = ", e^2, "\n",
      "7. Initial sample size (n₀): ", (z^2 * p * (1 - p)), " / ", e^2, " = ", res$initial, "\n"
    )
    
    if (res$finite_correction_applied) {
      calculation_steps <- paste0(calculation_steps,
        "\nFinite Population Correction Applied (N = ", N, "):\n",
        "8. Correction formula: n = n₀ / (1 + (n₀ - 1)/N)\n",
        "9. Calculate denominator: 1 + (", res$initial, " - 1)/", N, " = ", 
        (1 + (res$initial - 1)/N), "\n",
        "10. Corrected sample size: ", res$initial, " / ", (1 + (res$initial - 1)/N), " = ", res$corrected, "\n"
      )
    } else {
      calculation_steps <- paste0(calculation_steps,
        "\nNo finite population correction applied (population size not specified or infinite).\n"
      )
    }
    
    calculation_steps
  })
  
  output$cochranSample <- renderText({
    res <- cochranSampleSize()
    if (res$finite_correction_applied) {
      paste("Initial Sample Size (n₀):", ceiling(res$initial), "\n",
            "Corrected Sample Size (n):", ceiling(res$corrected))
    } else {
      paste("Sample Size (n₀):", ceiling(res$initial))
    }
  })
  
  output$cochranInterpretation <- renderText({
    res <- cochranSampleSize()
    p <- input$p
    z <- input$z
    e <- input$e_c
    N <- input$N_c
    
    interpretation <- paste(
      "Assuming a population proportion of", p,
      ", a confidence level corresponding to Z =", z,
      " (", round(pnorm(z) * 200 - 100, 1), "% confidence),",
      " and a margin of error of", e,
      ", the required sample size is", ceiling(res$corrected), "."
    )
    
    if (res$finite_correction_applied) {
      interpretation <- paste0(interpretation,
        " This includes a finite population correction for N = ", N, "."
      )
    } else {
      interpretation <- paste0(interpretation,
        " No finite population correction was applied (assumes infinite population)."
      )
    }
    
    interpretation
  })
  
  observeEvent(input$runPower, {
    result <- switch(input$testType,
                     "Independent t-test" = pwr.t.test(d = input$effectSize, sig.level = input$alpha, power = input$power, type = "two.sample"),
                     "Paired t-test" = pwr.t.test(d = input$effectSize, sig.level = input$alpha, power = input$power, type = "paired"),
                     "One-sample t-test" = pwr.t.test(d = input$effectSize, sig.level = input$alpha, power = input$power, type = "one.sample"),
                     "One-Way ANOVA" = pwr.anova.test(f = input$effectSize, sig.level = input$alpha, power = input$power, k = 1),
                     "Two-Way ANOVA" = pwr.anova.test(f = input$effectSize, sig.level = input$alpha, power = input$power, k = 2),
                     "Proportion" = pwr.p.test(h = input$effectSize, sig.level = input$alpha, power = input$power),
                     "Correlation" = pwr.r.test(r = input$effectSize, sig.level = input$alpha, power = input$power),
                     "Chi-squared" = pwr.chisq.test(w = input$effectSize, sig.level = input$alpha, power = input$power, df = 1),
                     "Simple Linear Regression" = pwr.f2.test(u = 1, f2 = input$effectSize, sig.level = input$alpha, power = input$power),
                     "Multiple Linear Regression" = pwr.f2.test(u = input$predictors, f2 = input$effectSize, sig.level = input$alpha, power = input$power)
    )
    output$powerResult <- renderPrint({ result })
  })
  
  observeEvent(input$runDesc, {
    req(input$dataInput)
    lines <- unlist(strsplit(input$dataInput, "[\n\r]+"))
    
    if (length(lines) < 2) {
      output$descResult <- renderPrint({ "Please include header (Variable name) and at least one row of data." })
      return()
    }
    
    header <- lines[1]
    data_lines <- lines[-1]
    
    suppressWarnings({
      nums <- as.numeric(data_lines)
      is_numeric <- !is.na(nums)
    })
    
    if (all(is_numeric)) {
      values <- nums[is_numeric]
      n_miss <- sum(is.na(nums))
      desc <- list(
        Mean = mean(values),
        SD = sd(values),
        Min = min(values),
        Max = max(values),
        Mode = as.numeric(names(sort(table(values), decreasing = TRUE)[1])),
        Median = median(values),
        Skewness = skewness(values),
        Kurtosis = kurtosis(values),
        Missing = n_miss,
        Shapiro_Wilk_p = shapiro.test(values)$p.value
      )
      desc$total <- sum(values)
      
      output$descResult <- renderPrint({
        cat(paste("Descriptive Statistics for:", header, "\n"))
        cat("(Numeric Data)\n\n")
        print(desc)
      })
    } else {
      categories <- trimws(data_lines)
      tbl <- table(factor(categories))
      cum_freq <- cumsum(tbl)
      percent <- prop.table(tbl) * 100
      n_miss <- sum(is.na(categories))
      
      df <- data.frame(
        Category = names(tbl),
        Frequency = as.vector(tbl),
        Percentage = round(as.vector(percent), 2),
        Cumulative_Frequency = as.vector(cum_freq)
      )
      
      df <- rbind(df, data.frame(
        Category = "Total", 
        Frequency = sum(tbl), 
        Percentage = 100, 
        Cumulative_Frequency = max(cum_freq)
      ))
      
      output$descResult <- renderPrint({
        cat(paste("Descriptive Statistics for:", header, "\n"))
        cat("(Categorical Data)\n\n")
        print(df)
        cat("\nMissing Values:", n_miss, "\n")
      })
    }
  })
}

shinyApp(ui = ui, server = server)
