library(shiny)
library(shinydashboard)
library(haven)
library(readxl)
library(foreign)
library(psych)
library(GPArotation)
library(ggplot2)
library(shinythemes)
library(DT)
library(readr)
library(plotly)
library(dplyr)
library(shinyjs)

ui <- fluidPage(
  useShinyjs(),
  # Add theme toggle script
  tags$script(HTML("
    $(document).ready(function() {
      // Check for saved theme preference
      const savedTheme = localStorage.getItem('factorguard-theme');
      if (savedTheme === 'dark') {
        $('body').addClass('dark-theme');
        $('#themeToggle i').removeClass('fa-sun').addClass('fa-moon');
      }
      
      // Theme toggle click handler
      $('#themeToggle').click(function() {
        $('body').toggleClass('dark-theme');
        const isDark = $('body').hasClass('dark-theme');
        if (isDark) {
          localStorage.setItem('factorguard-theme', 'dark');
          $('#themeToggle i').removeClass('fa-sun').addClass('fa-moon');
        } else {
          localStorage.setItem('factorguard-theme', 'light');
          $('#themeToggle i').removeClass('fa-moon').addClass('fa-sun');
        }
      });
    });
  ")),
  
  tags$head(
    tags$title("FactorGuard - Parallel Analysis Tool"),
    tags$style(HTML("
      /* Light theme (default) */
      :root {
        --bg-primary: #ffffff;
        --bg-secondary: #f8f9fa;
        --bg-tertiary: #e9ecef;
        --text-primary: #2d3748;
        --text-secondary: #4a5568;
        --text-tertiary: #718096;
        --accent-primary: #38b2ac;
        --accent-secondary: #2d9cdb;
        --border-color: #e2e8f0;
        --card-shadow: 0 20px 60px rgba(0, 0, 0, 0.1), 0 5px 15px rgba(0, 0, 0, 0.05);
        --card-shadow-hover: 0 25px 70px rgba(0, 0, 0, 0.15), 0 8px 20px rgba(0, 0, 0, 0.08);
        --gradient-primary: linear-gradient(135deg, #38b2ac 0%, #2d9cdb 100%);
        --gradient-secondary: linear-gradient(135deg, #ffffff 0%, #fdfdfd 100%);
        --gradient-dark: linear-gradient(135deg, #2d3748 0%, #1a202c 100%);
        --success-gradient: linear-gradient(135deg, #48bb78 0%, #38a169 100%);
        --warning-gradient: linear-gradient(135deg, #ed8936 0%, #dd6b20 100%);
        --info-gradient: linear-gradient(135deg, #4299e1 0%, #3182ce 100%);
      }
      
      /* Dark theme */
      .dark-theme {
        --bg-primary: #1a202c;
        --bg-secondary: #2d3748;
        --bg-tertiary: #4a5568;
        --text-primary: #f7fafc;
        --text-secondary: #e2e8f0;
        --text-tertiary: #cbd5e0;
        --accent-primary: #4fd1c7;
        --accent-secondary: #63b3ed;
        --border-color: #4a5568;
        --card-shadow: 0 20px 60px rgba(0, 0, 0, 0.3), 0 5px 15px rgba(0, 0, 0, 0.2);
        --card-shadow-hover: 0 25px 70px rgba(0, 0, 0, 0.4), 0 8px 20px rgba(0, 0, 0, 0.3);
        --gradient-primary: linear-gradient(135deg, #4fd1c7 0%, #63b3ed 100%);
        --gradient-secondary: linear-gradient(135deg, #2d3748 0%, #1a202c 100%);
        --gradient-dark: linear-gradient(135deg, #1a202c 0%, #0d1117 100%);
        --success-gradient: linear-gradient(135deg, #68d391 0%, #48bb78 100%);
        --warning-gradient: linear-gradient(135deg, #f6ad55 0%, #ed8936 100%);
        --info-gradient: linear-gradient(135deg, #63b3ed 0%, #4299e1 100%);
      }
      
      /* Theme transition */
      body {
        transition: background-color 0.3s ease, color 0.3s ease;
      }
      
      .main-container, 
      .sidebar-panel,
      .main-panel,
      .plot-container,
      .info-box,
      .formula-box,
      .result-box,
      .about-container,
      .file-input,
      .download-container,
      .syntax-help,
      .privacy-notice,
      .citation-box,
      .stat-card {
        transition: all 0.3s ease;
      }
      
      /* Rest of the styles with CSS variables */
      body {
        background: linear-gradient(135deg, var(--bg-tertiary) 0%, var(--bg-secondary) 100%);
        min-height: 100vh;
        margin: 0;
        padding: 20px;
        color: var(--text-primary);
      }
      
      .main-container {
        background: var(--gradient-secondary);
        border-radius: 20px;
        box-shadow: var(--card-shadow);
        padding: 40px;
        min-height: 95vh;
        display: flex;
        flex-direction: column;
        position: relative;
        overflow: hidden;
        border: 1px solid var(--border-color);
      }
      
      .main-container::before {
        content: '';
        position: absolute;
        top: 0;
        right: 0;
        width: 300px;
        height: 300px;
        background: linear-gradient(135deg, rgba(56, 178, 172, 0.05) 0%, rgba(45, 156, 219, 0.05) 100%);
        border-radius: 50%;
        transform: translate(100px, -100px);
      }
      
      .main-container::after {
        content: '';
        position: absolute;
        bottom: 0;
        left: 0;
        width: 200px;
        height: 200px;
        background: linear-gradient(135deg, rgba(56, 178, 172, 0.03) 0%, rgba(45, 156, 219, 0.03) 100%);
        border-radius: 50%;
        transform: translate(-50px, 50px);
      }
      
      .dark-theme .main-container::before,
      .dark-theme .main-container::after {
        background: linear-gradient(135deg, rgba(79, 209, 199, 0.05) 0%, rgba(99, 179, 237, 0.05) 100%);
      }
      
      @keyframes fadeInUp {
        from {
          opacity: 0;
          transform: translateY(30px);
        }
        to {
          opacity: 1;
          transform: translateY(0);
        }
      }
      
      @keyframes slideInLeft {
        from {
          opacity: 0;
          transform: translateX(-30px);
        }
        to {
          opacity: 1;
          transform: translateX(0);
        }
      }
      
      .title-container {
        text-align: center;
        margin-bottom: 40px;
        position: relative;
        z-index: 2;
      }
      
      .title-animation {
        animation: slideInLeft 1s ease-out;
        font-family: 'Montserrat', sans-serif;
        font-weight: 700;
        margin-bottom: 10px;
        background: var(--gradient-dark);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        font-size: 2.8em;
        letter-spacing: -0.5px;
      }
      
      .dark-theme .title-animation {
        background: linear-gradient(135deg, #f7fafc 0%, #e2e8f0 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
      }
      
      .subtitle-animation {
        animation: fadeInUp 1.2s ease-out 0.3s both;
        color: var(--text-tertiary);
        margin-bottom: 5px;
        font-weight: 500;
        font-size: 1.4em;
        font-family: 'Montserrat', sans-serif;
      }
      
      .tagline {
        animation: fadeInUp 1.2s ease-out 0.6s both;
        color: var(--text-tertiary);
        font-size: 1em;
        font-weight: 400;
        margin-top: 10px;
      }
      
      /* Theme Toggle Button */
      .theme-toggle-container {
        position: absolute;
        top: 20px;
        right: 20px;
        z-index: 1000;
      }
      
      .theme-toggle-btn {
        background: var(--gradient-primary);
        border: none;
        border-radius: 50%;
        width: 50px;
        height: 50px;
        display: flex;
        align-items: center;
        justify-content: center;
        cursor: pointer;
        box-shadow: 0 4px 15px rgba(56, 178, 172, 0.3);
        transition: all 0.3s ease;
        color: white;
        font-size: 1.2em;
      }
      
      .theme-toggle-btn:hover {
        transform: translateY(-3px);
        box-shadow: 0 8px 25px rgba(56, 178, 172, 0.4);
      }
      
      .dark-theme .theme-toggle-btn {
        box-shadow: 0 4px 15px rgba(79, 209, 199, 0.3);
      }
      
      .dark-theme .theme-toggle-btn:hover {
        box-shadow: 0 8px 25px rgba(79, 209, 199, 0.4);
      }
      
      .content-wrapper {
        flex: 1;
        display: flex;
        flex-direction: column;
        min-height: calc(100vh - 200px);
        position: relative;
        z-index: 2;
      }
      
      .main-content-area {
        flex: 1;
        margin-bottom: 30px;
      }
      
      .footer {
        background: var(--gradient-dark);
        color: var(--text-secondary);
        text-align: center;
        padding: 25px;
        border-radius: 15px;
        margin-top: auto;
        font-size: 0.9em;
        width: 100%;
        position: relative;
        z-index: 2;
        box-shadow: 0 -5px 20px rgba(0,0,0,0.05);
      }
      
      .footer::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        height: 4px;
        background: var(--gradient-primary);
        border-radius: 15px 15px 0 0;
      }
      
      /* Sidebar panel styling */
      .sidebar-panel {
        background: var(--bg-primary);
        padding: 35px;
        border-radius: 15px;
        box-shadow: var(--card-shadow);
        height: auto;
        min-height: 600px;
        overflow-y: auto;
        border: 1px solid var(--border-color);
        position: relative;
        z-index: 2;
      }
      
      .sidebar-panel::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        height: 4px;
        background: var(--gradient-primary);
        border-radius: 15px 15px 0 0;
      }
      
      .sidebar-title {
        font-family: 'Montserrat', sans-serif;
        font-weight: 600;
        color: var(--text-primary);
        margin-bottom: 30px;
        padding-bottom: 15px;
        border-bottom: 2px solid var(--border-color);
        font-size: 1.5em;
      }
      
      /* Main panel */
      .main-panel {
        padding: 0 35px;
        height: auto;
        min-height: 600px;
        overflow-y: auto;
        position: relative;
        z-index: 2;
      }
      
      /* Plot container */
      .plot-container {
        border: 1px solid var(--border-color);
        border-radius: 15px;
        padding: 25px;
        margin: 20px 0;
        box-shadow: var(--card-shadow);
        background: var(--bg-primary);
        min-height: 550px;
        width: 100%;
        position: relative;
        animation: fadeInUp 0.8s ease-out;
      }
      
      .plot-container::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        height: 4px;
        background: var(--gradient-primary);
        border-radius: 15px 15px 0 0;
      }
      
      /* Info boxes styling */
      .info-box {
        background: linear-gradient(135deg, var(--bg-secondary) 0%, var(--bg-tertiary) 100%);
        padding: 25px;
        border-radius: 12px;
        margin: 25px 0;
        border-left: 5px solid #ed8936;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
        animation: fadeInUp 0.8s ease-out;
        color: var(--text-primary);
      }
      
      .formula-box {
        background: linear-gradient(135deg, var(--bg-secondary) 0%, var(--bg-tertiary) 100%);
        padding: 25px;
        border-radius: 12px;
        margin: 25px 0;
        border-left: 5px solid #4299e1;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
        animation: fadeInUp 0.8s ease-out;
        color: var(--text-primary);
      }
      
      .result-box {
        background: linear-gradient(135deg, var(--bg-secondary) 0%, var(--bg-tertiary) 100%);
        padding: 25px;
        border-radius: 12px;
        margin: 25px 0;
        border-left: 5px solid #48bb78;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
        min-height: 200px;
        overflow-y: auto;
        animation: fadeInUp 0.8s ease-out;
        color: var(--text-primary);
      }
      
      .developer-info {
        background: var(--gradient-dark);
        color: white;
        padding: 25px;
        border-radius: 12px;
        margin: 25px 0;
        text-align: center;
        animation: fadeInUp 1s ease-out;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        position: relative;
        overflow: hidden;
      }
      
      .developer-info::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        height: 4px;
        background: var(--gradient-primary);
      }
      
      .developer-info h5 {
        font-family: 'Montserrat', sans-serif;
        font-weight: 600;
        margin-bottom: 15px;
        color: var(--text-secondary);
      }
      
      /* Button styling */
      .btn-primary {
        background: var(--gradient-primary);
        border: none;
        border-radius: 10px;
        font-weight: 600;
        padding: 14px 28px;
        transition: all 0.3s ease;
        color: white;
        width: 100%;
        margin: 15px 0;
        font-family: 'Montserrat', sans-serif;
        letter-spacing: 0.3px;
        box-shadow: 0 4px 15px rgba(56, 178, 172, 0.3);
      }
      
      .btn-primary:hover {
        transform: translateY(-3px);
        box-shadow: 0 8px 25px rgba(56, 178, 172, 0.4);
        color: white;
      }
      
      .dark-theme .btn-primary {
        box-shadow: 0 4px 15px rgba(79, 209, 199, 0.3);
      }
      
      .dark-theme .btn-primary:hover {
        box-shadow: 0 8px 25px rgba(79, 209, 199, 0.4);
      }
      
      .btn-success {
        background: var(--success-gradient);
        border: none;
        border-radius: 10px;
        font-weight: 600;
        padding: 12px 24px;
        transition: all 0.3s ease;
        color: white;
        width: 100%;
        box-shadow: 0 4px 15px rgba(72, 187, 120, 0.3);
      }
      
      .btn-info {
        background: var(--info-gradient);
        border: none;
        border-radius: 10px;
        font-weight: 600;
        padding: 12px 24px;
        transition: all 0.3s ease;
        color: white;
        width: 100%;
        box-shadow: 0 4px 15px rgba(66, 153, 225, 0.3);
      }
      
      /* Input styling */
      .form-control, .selectize-input {
        border-radius: 10px !important;
        border: 2px solid var(--border-color) !important;
        transition: all 0.3s ease;
        padding: 12px 16px;
        background-color: var(--bg-primary);
        color: var(--text-primary);
        font-size: 14px;
      }
      
      .form-control:focus, .selectize-input.focus {
        border-color: var(--accent-primary) !important;
        box-shadow: 0 0 0 3px rgba(56, 178, 172, 0.1) !important;
        outline: none;
      }
      
      .dark-theme .form-control:focus, 
      .dark-theme .selectize-input.focus {
        box-shadow: 0 0 0 3px rgba(79, 209, 199, 0.1) !important;
      }
      
      /* File input styling */
      .file-input {
        background: linear-gradient(135deg, var(--bg-secondary) 0%, var(--bg-tertiary) 100%);
        border-radius: 12px;
        border: 2px dashed var(--border-color);
        transition: all 0.3s ease;
        padding: 25px;
        margin: 20px 0;
        text-align: center;
        animation: fadeInUp 0.8s ease-out;
      }
      
      .file-input:hover {
        border-color: var(--accent-primary);
        background: linear-gradient(135deg, var(--bg-secondary) 0%, var(--bg-tertiary) 100%);
      }
      
      /* Tab styling */
      .nav-tabs {
        border-bottom: 2px solid var(--border-color);
        margin-bottom: 25px;
      }
      
      .nav-tabs > li > a {
        color: var(--text-tertiary);
        font-weight: 600;
        border-radius: 10px 10px 0 0;
        margin-right: 10px;
        padding: 12px 24px;
        font-family: 'Montserrat', sans-serif;
        border: 2px solid transparent;
        transition: all 0.3s ease;
        background: var(--bg-secondary);
      }
      
      .nav-tabs > li > a:hover {
        color: var(--accent-primary);
        background-color: var(--bg-tertiary);
        border-color: var(--border-color);
      }
      
      .nav-tabs > li.active > a {
        background: var(--bg-primary) !important;
        color: var(--text-primary) !important;
        border: 2px solid var(--border-color) !important;
        border-bottom: 2px solid var(--bg-primary) !important;
        position: relative;
        font-weight: 700;
      }
      
      .nav-tabs > li.active > a::after {
        content: '';
        position: absolute;
        bottom: -2px;
        left: 0;
        right: 0;
        height: 3px;
        background: var(--gradient-primary);
      }
      
      .download-container {
        text-align: center;
        margin: 30px 0;
        padding: 20px;
        background: linear-gradient(135deg, var(--bg-secondary) 0%, var(--bg-tertiary) 100%);
        border-radius: 12px;
        animation: fadeInUp 0.8s ease-out;
        border: 1px solid var(--border-color);
      }
      
      .plot-download-btn {
        background: var(--success-gradient);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 12px 28px;
        margin: 8px;
        font-weight: 600;
        transition: all 0.3s ease;
        font-family: 'Montserrat', sans-serif;
        letter-spacing: 0.3px;
        box-shadow: 0 4px 15px rgba(72, 187, 120, 0.3);
      }
      
      .plot-download-btn:hover {
        transform: translateY(-3px);
        box-shadow: 0 8px 25px rgba(72, 187, 120, 0.4);
        color: white;
      }
      
      /* Syntax help */
      .syntax-help {
        background: linear-gradient(135deg, var(--bg-secondary) 0%, var(--bg-tertiary) 100%);
        padding: 20px;
        margin: 25px 0;
        border-radius: 12px;
        border-left: 5px solid var(--accent-primary);
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
        animation: fadeInUp 0.8s ease-out;
        border: 1px solid var(--border-color);
      }
      
      .syntax-help h5 {
        font-family: 'Montserrat', sans-serif;
        font-weight: 600;
        color: var(--text-primary);
        margin-bottom: 15px;
      }
      
      .syntax-help ul {
        padding-left: 20px;
      }
      
      .syntax-help li {
        margin-bottom: 8px;
        color: var(--text-secondary);
      }
      
      /* Privacy notice */
      .privacy-notice {
        background: linear-gradient(135deg, var(--bg-secondary) 0%, var(--bg-tertiary) 100%);
        border: 2px solid #ed8936;
        color: var(--text-primary);
        padding: 20px;
        border-radius: 12px;
        margin: 25px 0;
        font-size: 14px;
        animation: fadeInUp 0.8s ease-out;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
      }
      
      .privacy-notice strong {
        color: #c05621;
      }
      
      /* About page styling */
      .about-container {
        background: var(--bg-primary);
        padding: 40px;
        border-radius: 20px;
        box-shadow: var(--card-shadow);
        margin: 25px 0;
        overflow-y: auto;
        min-height: 600px;
        border: 1px solid var(--border-color);
        animation: fadeInUp 0.8s ease-out;
        position: relative;
      }
      
      .about-container::before {
        content: '';
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        height: 4px;
        background: var(--gradient-primary);
        border-radius: 20px 20px 0 0;
      }
      
      .about-section {
        margin-bottom: 35px;
        padding-bottom: 25px;
        border-bottom: 2px solid var(--border-color);
      }
      
      .about-section:last-child {
        border-bottom: none;
      }
      
      .about-section h3 {
        font-family: 'Montserrat', sans-serif;
        font-weight: 700;
        color: var(--text-primary);
        margin-bottom: 20px;
        font-size: 1.8em;
      }
      
      .about-section h4 {
        font-family: 'Montserrat', sans-serif;
        font-weight: 600;
        color: var(--text-secondary);
        margin-bottom: 15px;
      }
      
      .citation-box {
        background: linear-gradient(135deg, var(--bg-secondary) 0%, var(--bg-tertiary) 100%);
        padding: 20px;
        margin: 20px 0;
        border-radius: 10px;
        border-left: 5px solid #4299e1;
        font-size: 13px;
        line-height: 1.6;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
        color: var(--text-primary);
      }
      
      .feature-list {
        list-style-type: none;
        padding-left: 0;
      }
      
      .feature-list li {
        padding: 12px 0;
        padding-left: 35px;
        position: relative;
        color: var(--text-secondary);
        font-size: 15px;
      }
      
      .feature-list li:before {
        content: '✓';
        position: absolute;
        left: 0;
        color: var(--accent-primary);
        font-weight: bold;
        font-size: 16px;
        width: 24px;
        height: 24px;
        background: var(--bg-tertiary);
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
      }
      
      /* Custom scrollbar */
      ::-webkit-scrollbar {
        width: 10px;
      }
      
      ::-webkit-scrollbar-track {
        background: var(--bg-secondary);
        border-radius: 5px;
      }
      
      ::-webkit-scrollbar-thumb {
        background: var(--gradient-primary);
        border-radius: 5px;
      }
      
      ::-webkit-scrollbar-thumb:hover {
        background: var(--gradient-primary);
      }
      
      /* Label styling */
      .control-label {
        font-weight: 600;
        color: var(--text-secondary);
        margin-bottom: 8px;
        font-family: 'Montserrat', sans-serif;
      }
      
      /* Statistic cards */
      .stat-card {
        background: var(--bg-primary);
        border-radius: 12px;
        padding: 20px;
        margin: 10px 0;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
        border: 1px solid var(--border-color);
        text-align: center;
      }
      
      .stat-card h4 {
        color: var(--accent-primary);
        font-family: 'Montserrat', sans-serif;
        margin-bottom: 10px;
      }
      
      .stat-card p {
        color: var(--text-tertiary);
        font-size: 0.9em;
      }
      
      /* Plotly dark theme adjustments */
      .dark-theme .js-plotly-plot .plotly {
        background-color: transparent !important;
      }
      
      .dark-theme .modebar {
        background-color: var(--bg-primary) !important;
      }
      
      .dark-theme .modebar-btn path {
        fill: var(--text-secondary) !important;
      }
      
      .dark-theme .modebar-btn:hover path {
        fill: var(--accent-primary) !important;
      }
      
      /* Selectize dropdown dark theme */
      .dark-theme .selectize-dropdown {
        background-color: var(--bg-primary);
        border: 1px solid var(--border-color);
        color: var(--text-primary);
      }
      
      .dark-theme .selectize-dropdown .active {
        background-color: var(--bg-tertiary);
        color: var(--text-primary);
      }
      
      .dark-theme .selectize-dropdown .selected {
        background-color: var(--accent-primary);
        color: white;
      }
      
      .dark-theme .selectize-input {
        background-color: var(--bg-primary) !important;
        color: var(--text-primary) !important;
      }
      
      .dark-theme .selectize-input input {
        color: var(--text-primary) !important;
      }
      
      /* Datatable dark theme */
      .dark-theme .dataTables_wrapper {
        color: var(--text-primary);
      }
      
      .dark-theme .dataTables_filter input {
        background-color: var(--bg-primary);
        color: var(--text-primary);
        border: 1px solid var(--border-color);
      }
      
      .dark-theme .dataTables_info {
        color: var(--text-secondary) !important;
      }
    "))
  ),
  
  # Main container
  div(class = "main-container",
      
      # Theme Toggle Button
      div(class = "theme-toggle-container",
          actionButton("themeToggle", "",
                       class = "theme-toggle-btn",
                       icon = icon("sun"))
      ),
      
      # Application title with enhanced design
      div(class = "title-container",
          div(class = "title-animation",
              h1("FactorGuard - A factor-retention decision tool for EFA")
          )
      ),
      
      # Main content wrapper
      div(class = "content-wrapper",
          div(class = "main-content-area",
              sidebarLayout(
                sidebarPanel(
                  width = 4,
                  class = "sidebar-panel",
                  h3("Data Input & Settings", class = "sidebar-title"),
                  
                  # Privacy notice
                  div(class = "privacy-notice",
                      icon("shield-alt", style = "color: #ed8936; margin-right: 10px;"),
                      strong("Privacy First:"),
                      " All data processing occurs locally in your browser. No data is stored on servers."
                  ),
                  
                  # File upload
                  div(class = "file-input",
                      fileInput("file_parallel", "Upload Dataset", 
                                accept = c(".csv", ".xlsx", ".xls", ".sav", ".dta"),
                                width = "100%",
                                buttonLabel = "Browse Files...",
                                placeholder = "CSV, Excel, SPSS, or Stata files")
                  ),
                  
                  # Variable selection
                  div(class = "form-group",
                      tags$label("Select Variables:", class = "control-label"),
                      uiOutput("varSelect_parallel")
                  ),
                  
                  # Number of simulations
                  div(class = "form-group",
                      tags$label("Number of Simulations:", class = "control-label"),
                      numericInput("n_iter", NULL, 
                                   value = 20, min = 10, max = 1000, 
                                   step = 10, width = "100%")
                  ),
                  
                  # Action button
                  actionButton("run_parallel", "Run Parallel Analysis", 
                               class = "btn-primary", icon = icon("play-circle")),
                  
                  br(), br(),
                  
                  # Syntax help
                  div(class = "syntax-help",
                      h5("📊 Technical Implementation:"),
                      tags$ul(
                        tags$li("Uses ", code("fa.parallel()"), " from psych package for analysis"),
                        tags$li("Generates interactive scree plots using plotly package"),
                        tags$li("Compares actual vs. random data eigenvalues"),
                        tags$li("More robust than traditional eigenvalue rules"),
                        tags$li("Monte Carlo simulation-based approach"),
                        tags$li("Results appear in the analysis tabs")
                      )
                  ),
                  
                  # Developer information
                  div(class = "developer-info",
                      h5("👨‍💻 Developer Information"),
                      p(icon("user-md", style = "margin-right: 10px;"), "Mudasir Mohammed Ibrahim, RN"),
                      p(icon("envelope", style = "margin-right: 10px;"), "mudassiribrahim30@gmail.com"),
                      p(icon("globe", style = "margin-right: 10px;"), "www.mudasiribrahim.com")
                  )
                ),
                
                mainPanel(
                  width = 8,
                  class = "main-panel",
                  tabsetPanel(
                    tabPanel("📈 Parallel Analysis",
                             div(class = "plot-container",
                                 plotOutput("faParallelPlot", height = "500px")
                             ),
                             div(class = "download-container",
                                 h5("Export Results", style = "margin-bottom: 15px; color: var(--text-secondary);"),
                                 downloadButton("downloadParallel", "Download Parallel Plot", 
                                                class = "plot-download-btn")
                             ),
                             div(class = "result-box",
                                 h4("Analysis Results", style = "margin-bottom: 15px; color: var(--text-primary); font-family: 'Montserrat';"),
                                 verbatimTextOutput("parallelInterpretation")
                             )
                    ),
                    
                    tabPanel("📊 Scree Plot",
                             div(class = "plot-container",
                                 plotlyOutput("screePlot", height = "500px")
                             )
                    ),
                    
                    tabPanel("ℹ️ About FactorGuard",
                             div(class = "about-container",
                                 div(class = "about-section",
                                     h3("👤 About the Developer"),
                                     p("My name is ", strong("Mudasir Mohammed Ibrahim"), ", a Registered Nurse with a strong passion for healthcare research and data analysis."),
                                     p("Through extensive clinical and research experience, I identified a persistent methodological challenge: many healthcare professionals and researchers struggle to determine the appropriate number of factors to extract in exploratory factor analysis, particularly due to ambiguities in interpreting scree plots and eigenvalues. These challenges often undermine the validity and interpretability of factor analytic results. This realization motivated the development of FactorGuard—a sophisticated yet user-friendly application designed to support transparent, data-driven factor retention decisions and make factor analysis more accessible and reliable for researchers in the healthcare and social sciences.")
                                 ),
                                 
                                 div(class = "about-section",
                                     h3("⚙️ Application Features"),
                                     tags$ul(class = "feature-list",
                                             tags$li("Parallel Analysis: Utilizes ", code("fa.parallel()"), " function from the psych R package for robust factor retention decisions"),
                                             tags$li("Interactive Scree Plots: Professional visualizations generated using plotly package"),
                                             tags$li("Multi-format Data Support: Import from CSV, Excel, SPSS, and Stata formats"),
                                             tags$li("Publication-Ready Exports: High-quality downloadable plots for research papers"),
                                             tags$li("Interactive Exploration: Hover-enabled plots for detailed eigenvalue examination")
                                     )
                                 ),
                                 
                                 div(class = "about-section",
                                     h3("🔬 Methodological Foundation"),
                                     p("FactorGuard implements ", strong("parallel analysis"), " - a statistically rigorous method based on:"),
                                     p("1. ", strong("Horn's Parallel Analysis (1965):"), " Uses ", code("fa.parallel()"), " to compare actual eigenvalues with eigenvalues from random data matrices. Factors are retained when actual eigenvalues exceed random data percentiles."),
                                     p("2. ", strong("Robust Comparison:"), " Addresses limitations of traditional methods (Kaiser's criterion, scree plot interpretation) with empirical, simulation-based evidence."),
                                     p("3. ", strong("Monte Carlo Simulations:"), " Generates random datasets matching original data dimensions for appropriate benchmarks."),
                                     p("4. ", strong("Psych Package Integration:"), " Leverages well-established R functions for reliable, validated results.")
                                 ),
                                 
                                 div(class = "about-section",
                                     h3("🔒 Data Privacy & Security"),
                                     h4(icon("shield-alt"), " Local Processing Guarantee"),
                                     p("FactorGuard prioritizes your research data privacy:"),
                                     tags$ul(class = "feature-list",
                                             tags$li("All computations occur locally in your web browser"),
                                             tags$li("No data transmission to external servers"),
                                             tags$li("No user data storage or collection"),
                                             tags$li("No tracking or analytics implementation"),
                                             tags$li("Complete data sovereignty maintained")
                                     ),
                                     p("This approach ensures maximum confidentiality for sensitive research data, particularly crucial in healthcare and social science research.")
                                 ),
                                 
                                 div(class = "about-section",
                                     h3("📚 References & Citations"),
                                     div(class = "citation-box",
                                         p(strong("Primary Methodological Reference:")),
                                         p("Horn, J. L. (1965). A Rationale and Test for the Number of Factors in Factor Analysis. Psychometrika, 30, 179-185.")
                                     ),
                                     
                                     div(class = "citation-box",
                                         p(strong("R Package Reference:")),
                                         p("Revelle, W. (2013). psych: Procedures for Psychological, Psychometric, and Personality Research. Northwestern University.")
                                     ),
                                     
                                     div(class = "citation-box",
                                         p(strong("Application Citation:")),
                                         p("When using FactorGuard in research, please cite:"),
                                         p("Ibrahim, M. M. (2026). FactorGuard: A factor-retention decision tool for EFA [Computer software].")
                                     )
                                 ),
                                 
                                 div(class = "about-section",
                                     h3("💻 Technical Specifications"),
                                     div(class = "stat-card",
                                         h4("Version 1.0.0"),
                                         p("Current Release")
                                     ),
                                     div(class = "stat-card",
                                         h4("Built With R/Shiny"),
                                         p("psych, ggplot2, plotly")
                                     ),
                                     div(class = "stat-card",
                                         h4("Core Functions"),
                                         p("fa.parallel() & plotly()")
                                     ),
                                     br(),
                                     p(strong("Release Date:"), " 22nd January 2026"),
                                     p(strong("Platform:"), " Web-based Shiny Application"),
                                     p(strong("Compatibility:"), " All modern browsers with JavaScript"),
                                     p(strong("License:"), " CC BY 4.0")
                                 )
                             )
                    )
                  )
                )
              )
          ),
          
          # Professional Footer
          div(class = "footer",
              div(style = "display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap;",
                  div(style = "text-align: left;",
                      h5("FactorGuard", style = "margin: 0; color: white; font-family: 'Montserrat';"),
                      p("A factor-retention decision tool for EFA", style = "margin: 5px 0 0 0; font-size: 0.85em; color: white; opacity: 0.8;")
                  ),
                  div(style = "text-align: center; flex-grow: 1;",
                      p("Built with ❤️ in Ghana by Mudasir Mohammed Ibrahim, RN", style = "margin: 0; color: white;"),
                      p("© 2026 CC BY 4.0 | Data processed locally - Never stored", 
                        style = "margin: 5px 0 0 0; font-size: 0.85em; color: white; opacity: 0.8;")
                  ),
                  div(style = "text-align: right;",
                      p(icon("r-project"), " R/Shiny Application", style = "margin: 0; color: white;"),
                      p("Version 1.0.0", style = "margin: 5px 0 0 0; font-size: 0.85em; color: white; opacity: 0.8;")
                  )
              )
          )
      )
  )
)

server <- function(input, output, session) {
  # Shared data reactive
  data_shared <- reactive({
    if (!is.null(input$file_parallel)) {
      file <- input$file_parallel
    } else {
      return(NULL)
    }
    
    req(file)
    ext <- tools::file_ext(file$name)
    switch(ext,
           csv = read_csv(file$datapath),
           xlsx = read_excel(file$datapath),
           xls = read_excel(file$datapath),
           sav = read_sav(file$datapath),
           dta = read_dta(file$datapath),
           validate("Unsupported file type")
    )
  })
  
  # Variable selection UI
  output$varSelect_parallel <- renderUI({
    req(data_shared())
    selectInput("vars_parallel", NULL, choices = names(data_shared()), multiple = TRUE,
                selectize = TRUE)
  })
  
  # Selected data
  selectedData_parallel <- reactive({
    req(input$vars_parallel)
    data_shared()[, input$vars_parallel, drop = FALSE]
  })
  
  # Reactive values for storing results
  parallel_result <- reactiveVal(NULL)
  
  # Run parallel analysis
  observeEvent(input$run_parallel, {
    req(selectedData_parallel())
    
    # Show notification
    showNotification("Running parallel analysis...", type = "message", duration = 3)
    
    # Run parallel analysis using psych package functions
    tryCatch({
      parallel <- fa.parallel(
        selectedData_parallel(),
        fa = "fa",
        n.iter = input$n_iter,
        plot = FALSE,
        show.legend = FALSE
      )
      
      parallel_result(parallel)
      
      # Parallel analysis interpretation
      output$parallelInterpretation <- renderPrint({
        req(parallel_result())
        parallel <- parallel_result()
        cat("┌─────────────────────────────────────────────┐\n")
        cat("│      PARALLEL ANALYSIS RESULTS              │\n")
        cat("│     (using psych::fa.parallel)              │\n")
        cat("└─────────────────────────────────────────────┘\n\n")
        
        cat("📊 SUGGESTED FACTORS: ", parallel$nfact, "\n")
        cat("─────────────────────────────────────────────\n\n")
        
        cat("METHODOLOGY:\n")
        cat("────────────\n")
        cat("• Compares actual eigenvalues with random data eigenvalues\n")
        cat("• Uses Monte Carlo simulations (", input$n_iter, " iterations)\n")
        cat("• More robust than Kaiser's criterion (eigenvalue > 1)\n\n")
        
        cat("INTERPRETATION:\n")
        cat("───────────────\n")
        cat("1. Parallel analysis suggests retaining ", parallel$nfact, " factor(s)\n")
        cat("2. Factors retained when actual eigenvalues exceed\n")
        cat("   corresponding random data percentiles (red line)\n\n")
        
        cat("RECOMMENDATIONS:\n")
        cat("────────────────\n")
        if (parallel$nfact == 1) {
          cat("✓ Consider single-factor solution if conceptually justified\n")
          cat("✓ Verify with scree plot elbow point\n")
          cat("✓ Assess factor loadings and communalities\n")
        } else if (parallel$nfact > 1) {
          cat("✓ Extract ", parallel$nfact, " factors for further analysis\n")
          cat("✓ Compare with scree plot results\n")
          cat("✓ Evaluate factor interpretability and structure\n")
          cat("✓ Consider oblique rotation if factors correlate\n")
        } else {
          cat("⚠ No factors suggested by parallel analysis\n")
          cat("• Data may not be suitable for factor analysis\n")
          cat("• Check correlation matrix and sampling adequacy\n")
          cat("• Consider alternative dimensionality methods\n")
        }
        
        cat("\n─────────────────────────────────────────────\n")
        cat("Analysis performed using psych::fa.parallel()\n")
        cat("Generated: ", format(Sys.time(), "%Y-%m-%d %H:%M:%S"), "\n")
      })
      
      # Parallel analysis plot with professional styling
      output$faParallelPlot <- renderPlot({
        req(parallel_result())
        parallel <- parallel_result()
        
        # Create custom plot
        n_factors <- length(parallel$fa.values)
        plot_data <- data.frame(
          Factor = 1:n_factors,
          Actual = parallel$fa.values,
          Simulated = parallel$fa.simr
        )
        
        # Determine number of suggested factors
        n_suggested <- parallel$nfact
        
        # Professional plot with enhanced styling
        p <- ggplot(plot_data, aes(x = Factor)) +
          geom_line(aes(y = Actual, color = "Actual Data"), size = 1.8, alpha = 0.9) +
          geom_point(aes(y = Actual, color = "Actual Data"), size = 4, shape = 21, fill = "white", stroke = 1.5) +
          geom_line(aes(y = Simulated, color = "Simulated Data"), size = 1.5, linetype = "dashed", alpha = 0.8) +
          geom_point(aes(y = Simulated, color = "Simulated Data"), size = 3.5, shape = 24, fill = "white", stroke = 1.2) +
          geom_hline(yintercept = 1, linetype = "dotted", color = "#718096", size = 0.8) +
          scale_color_manual(values = c("Actual Data" = "#38b2ac", "Simulated Data" = "#e53e3e")) +
          labs(
            title = "Parallel Analysis Scree Plot",
            subtitle = paste("Suggested number of factors to retain:", n_suggested),
            x = "Factor Number",
            y = "Eigenvalue",
            color = "Data Type",
            caption = paste("Generated using psych::fa.parallel() |", 
                            format(Sys.time(), "%Y-%m-%d %H:%M"))
          ) +
          theme_minimal() +
          theme(
            plot.title = element_text(hjust = 0.5, size = 18, face = "bold", 
                                      color = "#2d3748", family = "Montserrat"),
            plot.subtitle = element_text(hjust = 0.5, size = 14, color = "#4a5568", 
                                         margin = margin(b = 15)),
            axis.title = element_text(size = 12, color = "#4a5568", face = "bold"),
            axis.text = element_text(color = "#718096"),
            legend.position = "bottom",
            legend.title = element_text(face = "bold", color = "#2d3748"),
            legend.text = element_text(color = "#4a5568"),
            panel.grid.major = element_line(color = "#e2e8f0", size = 0.5),
            panel.grid.minor = element_blank(),
            plot.margin = unit(c(20, 20, 20, 20), "points"),
            plot.background = element_rect(fill = "white", color = NA),
            panel.background = element_rect(fill = "white", color = NA),
            plot.caption = element_text(color = "#a0aec0", size = 9, hjust = 1)
          ) +
          annotate("text", x = max(plot_data$Factor), y = 1.05, 
                   label = "Eigenvalue = 1", hjust = 1, vjust = 0, 
                   color = "#718096", size = 3.5, family = "Open Sans") +
          guides(color = guide_legend(override.aes = list(size = 3)))
        
        # Add annotation for number of suggested factors
        if (n_suggested > 0 && n_suggested <= nrow(plot_data)) {
          p <- p + 
            annotate("rect", 
                     xmin = n_suggested - 0.3, xmax = n_suggested + 0.3,
                     ymin = plot_data$Actual[n_suggested] - 0.1, 
                     ymax = plot_data$Actual[n_suggested] + 0.1,
                     fill = "#48bb78", alpha = 0.15) +
            annotate("text", 
                     x = n_suggested, 
                     y = plot_data$Actual[n_suggested] + 0.25,
                     label = paste("Factor", n_suggested),
                     color = "#38a169", 
                     fontface = "bold",
                     size = 4.5,
                     family = "Montserrat")
        }
        
        # Add subtle watermark
        p <- p + 
          annotate("text", 
                   x = max(plot_data$Factor) * 0.98, 
                   y = max(plot_data$Actual) * 0.05,
                   label = "FactorGuard",
                   color = "#e2e8f0", 
                   size = 8,
                   hjust = 1,
                   vjust = 0,
                   alpha = 0.5,
                   family = "Montserrat",
                   fontface = "bold")
        
        p
      }, height = 500)
      
      # Professional Scree Plot
      output$screePlot <- renderPlotly({
        ev <- eigen(cor(selectedData_parallel(), use = "pairwise.complete.obs"))
        scree_data <- data.frame(
          Factor = 1:length(ev$values), 
          Eigenvalue = ev$values,
          Type = "Actual"
        )
        
        # Add parallel analysis line if available
        if (!is.null(parallel_result())) {
          parallel <- parallel_result()
          if (length(parallel$fa.simr) >= nrow(scree_data)) {
            scree_data$Simulated <- parallel$fa.simr[1:nrow(scree_data)]
          }
        }
        
        # Create professional plot
        p <- plot_ly(scree_data, x = ~Factor, y = ~Eigenvalue, 
                     type = 'scatter', mode = 'lines+markers',
                     name = 'Actual Eigenvalues',
                     marker = list(size = 10, color = '#38b2ac', 
                                   line = list(color = 'white', width = 1.5)),
                     line = list(color = '#38b2ac', width = 3),
                     hovertemplate = paste(
                       "<b>Factor:</b> %{x}<br>",
                       "<b>Eigenvalue:</b> %{y:.3f}<br>",
                       "<extra></extra>"
                     ))
        
        # Add simulated line if available
        if ("Simulated" %in% names(scree_data)) {
          p <- p %>% 
            add_trace(x = ~Factor, y = ~Simulated, 
                      type = 'scatter', mode = 'lines+markers',
                      name = 'Parallel Analysis',
                      line = list(color = '#e53e3e', width = 2.5, dash = 'dash'),
                      marker = list(size = 8, color = '#e53e3e', symbol = 'triangle-up',
                                    line = list(color = 'white', width = 1)),
                      hovertemplate = paste(
                        "<b>Factor:</b> %{x}<br>",
                        "<b>Simulated Eigenvalue:</b> %{y:.3f}<br>",
                        "<extra></extra>"
                      ))
        }
        
        # Add layout with professional styling
        p <- p %>% 
          layout(
            title = list(
              text = "<b>Scree Plot</b>",
              font = list(size = 20, family = "Montserrat", color = "#2d3748"),
              x = 0.05
            ),
            xaxis = list(
              title = list(text = "<b>Factor Number</b>", 
                           font = list(size = 14, color = "#4a5568")),
              tickmode = "linear",
              gridcolor = "#e2e8f0",
              tickfont = list(color = "#718096")
            ),
            yaxis = list(
              title = list(text = "<b>Eigenvalue</b>", 
                           font = list(size = 14, color = "#4a5568")),
              gridcolor = "#e2e8f0",
              tickfont = list(color = "#718096")
            ),
            hovermode = 'x unified',
            showlegend = TRUE,
            legend = list(
              orientation = 'h',
              y = -0.15,
              font = list(color = "#4a5568", size = 12),
              bgcolor = 'rgba(255,255,255,0.8)'
            ),
            margin = list(l = 60, r = 40, t = 60, b = 80),
            plot_bgcolor = 'white',
            paper_bgcolor = 'white',
            shapes = list(
              list(
                type = "line",
                x0 = 0,
                x1 = 1,
                xref = "paper",
                y0 = 1,
                y1 = 1,
                yref = "y",
                line = list(color = "#718096", dash = "dot", width = 1)
              )
            ),
            annotations = list(
              list(
                x = 1,
                xref = "paper",
                y = 1.05,
                yref = "y",
                text = "Eigenvalue = 1",
                showarrow = FALSE,
                font = list(color = "#718096", size = 11, family = "Open Sans"),
                xanchor = "right",
                yanchor = "bottom"
              ),
              list(
                x = 0.5,
                xref = "paper",
                y = -0.25,
                yref = "paper",
                text = paste("Generated using psych package functions |", 
                             format(Sys.time(), "%Y-%m-%d %H:%M")),
                showarrow = FALSE,
                font = list(color = "#a0aec0", size = 10, family = "Open Sans"),
                xanchor = "center"
              )
            )
          )
        
        p
      })
      
      # Download handler for parallel plot
      output$downloadParallel <- downloadHandler(
        filename = function() { 
          paste0("parallel_analysis_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".png") 
        },
        content = function(file) {
          parallel <- parallel_result()
          n_factors <- length(parallel$fa.values)
          plot_data <- data.frame(
            Factor = 1:n_factors,
            Actual = parallel$fa.values,
            Simulated = parallel$fa.simr
          )
          
          n_suggested <- parallel$nfact
          
          # High-quality export plot
          p <- ggplot(plot_data, aes(x = Factor)) +
            geom_line(aes(y = Actual, color = "Actual Data"), size = 2) +
            geom_point(aes(y = Actual, color = "Actual Data"), size = 4, shape = 21, fill = "white", stroke = 1.5) +
            geom_line(aes(y = Simulated, color = "Simulated Data"), size = 1.5, linetype = "dashed") +
            geom_point(aes(y = Simulated, color = "Simulated Data"), size = 3.5, shape = 24, fill = "white", stroke = 1.2) +
            geom_hline(yintercept = 1, linetype = "dotted", color = "#718096", size = 0.8) +
            scale_color_manual(values = c("Actual Data" = "#38b2ac", "Simulated Data" = "#e53e3e")) +
            labs(
              title = "Parallel Analysis Scree Plot",
              subtitle = paste("Suggested number of factors to retain:", n_suggested),
              x = "Factor Number",
              y = "Eigenvalue",
              color = "Data Type",
              caption = paste("Generated using psych::fa.parallel() | FactorGuard |", 
                              format(Sys.time(), "%Y-%m-%d %H:%M:%S"))
            ) +
            theme_minimal() +
            theme(
              plot.title = element_text(hjust = 0.5, size = 20, face = "bold", 
                                        color = "#2d3748", family = "Montserrat"),
              plot.subtitle = element_text(hjust = 0.5, size = 16, color = "#4a5568", 
                                           margin = margin(b = 20)),
              axis.title = element_text(size = 14, color = "#4a5568", face = "bold"),
              axis.text = element_text(size = 12, color = "#718096"),
              legend.position = "bottom",
              legend.title = element_text(face = "bold", size = 12, color = "#2d3748"),
              legend.text = element_text(size = 11, color = "#4a5568"),
              panel.grid.major = element_line(color = "#e2e8f0", size = 0.5),
              panel.grid.minor = element_blank(),
              plot.margin = unit(c(25, 25, 25, 25), "points"),
              plot.caption = element_text(color = "#a0aec0", size = 10, hjust = 1)
            )
          
          # Add annotation for suggested factors
          if (n_suggested > 0 && n_suggested <= nrow(plot_data)) {
            p <- p + 
              annotate("rect", 
                       xmin = n_suggested - 0.3, xmax = n_suggested + 0.3,
                       ymin = plot_data$Actual[n_suggested] - 0.1, 
                       ymax = plot_data$Actual[n_suggested] + 0.1,
                       fill = "#48bb78", alpha = 0.15) +
              annotate("text", 
                       x = n_suggested, 
                       y = plot_data$Actual[n_suggested] + 0.25,
                       label = paste("Factor", n_suggested),
                       color = "#38a169", 
                       fontface = "bold",
                       size = 5,
                       family = "Montserrat")
          }
          
          ggsave(file, plot = p, width = 12, height = 8, dpi = 300, bg = "white")
        }
      )
      
      # Download handler for scree plot
      output$downloadScree <- downloadHandler(
        filename = function() { 
          paste0("scree_plot_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".png") 
        },
        content = function(file) {
          ev <- eigen(cor(selectedData_parallel(), use = "pairwise.complete.obs"))
          scree_data <- data.frame(Factor = 1:length(ev$values), Eigenvalue = ev$values)
          
          # Get parallel results if available
          n_suggested <- if (!is.null(parallel_result())) parallel_result()$nfact else NA
          
          p <- ggplot(scree_data, aes(x = Factor, y = Eigenvalue)) +
            geom_point(size = 4, color = "#38b2ac", shape = 21, fill = "white", stroke = 1.5) + 
            geom_line(color = "#38b2ac", size = 2) +
            geom_hline(yintercept = 1, linetype = "dotted", color = "#718096", size = 0.8) +
            labs(
              title = "Scree Plot", 
              subtitle = if (!is.na(n_suggested)) paste("Parallel analysis suggests", n_suggested, "factors") else NULL,
              x = "Factor Number", 
              y = "Eigenvalue",
              caption = paste("Generated using psych package functions | FactorGuard |", 
                              format(Sys.time(), "%Y-%m-%d %H:%M:%S"))
            ) +
            theme_minimal() +
            theme(
              plot.title = element_text(hjust = 0.5, size = 20, face = "bold", 
                                        color = "#2d3748", family = "Montserrat"),
              plot.subtitle = element_text(hjust = 0.5, size = 14, color = "#4a5568"),
              plot.caption = element_text(hjust = 1, size = 10, color = "#a0aec0"),
              axis.title = element_text(size = 14, color = "#4a5568", face = "bold"),
              axis.text = element_text(size = 12, color = "#718096"),
              panel.grid.major = element_line(color = "#e2e8f0", size = 0.5),
              panel.grid.minor = element_blank(),
              plot.margin = unit(c(25, 25, 25, 25), "points")
            ) +
            annotate("text", x = max(scree_data$Factor), y = 1.05, 
                     label = "Eigenvalue = 1", hjust = 1, vjust = 0, 
                     color = "#718096", size = 4.5, family = "Open Sans")
          
          # Add factor annotation if parallel analysis was run
          if (!is.na(n_suggested) && n_suggested > 0 && n_suggested <= nrow(scree_data)) {
            p <- p + 
              annotate("rect", 
                       xmin = n_suggested - 0.3, xmax = n_suggested + 0.3,
                       ymin = scree_data$Eigenvalue[n_suggested] - 0.1, 
                       ymax = scree_data$Eigenvalue[n_suggested] + 0.1,
                       fill = "#48bb78", alpha = 0.15) +
              annotate("text", 
                       x = n_suggested, 
                       y = scree_data$Eigenvalue[n_suggested] + 0.25,
                       label = paste("Factor", n_suggested),
                       color = "#38a169", 
                       fontface = "bold",
                       size = 5,
                       family = "Montserrat")
          }
          
          ggsave(file, plot = p, width = 12, height = 8, dpi = 300, bg = "white")
        }
      )
      
      # Show notification 
      showNotification("Running parallel analysis...", type = "message", duration = 3)
      
    }, error = function(e) {
      showNotification(paste("❌ Error:", e$message), type = "error", duration = 6)
    })
  })
}

shinyApp(ui, server)
