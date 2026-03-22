# Fix Shinyapps.io timeout issue
options(shinyapps.http_timeout = 9000)   
options(timeout = 900)                  # General R timeout

library(curl)

safe_api_call <- function(url) {
  h <- new_handle()
  handle_setopt(h,
                timeout = 9000,         
                connecttimeout = 1800920     
  )
  curl_fetch_memory(url, handle = h)
}


res <- safe_api_call("https://api.shinyapps.io/your_endpoint")










