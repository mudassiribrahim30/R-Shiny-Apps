# Fix Shinyapps.io timeout issue
options(shinyapps.http_timeout = 900)   # 5 minutes
options(timeout = 900)                  # General R timeout

library(curl)

safe_api_call <- function(url) {
  h <- new_handle()
  handle_setopt(h,
                timeout = 900,          # 5 minutes
                connecttimeout = 120     # 1 minute
  )
  curl_fetch_memory(url, handle = h)
}


res <- safe_api_call("https://api.shinyapps.io/your_endpoint")











