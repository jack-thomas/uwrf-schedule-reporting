FROM rocker/r-ver:4.6.2

# system packages
RUN apt-get update && apt-get install -y --no-install-recommends \
    libcurl4-openssl-dev \
    libssl-dev \
    libxml2-dev \
    zlib1g-dev \
  && rm -rf /var/lib/apt/lists/*

RUN Rscript -e '\
  options(HTTPUserAgent = sprintf("R/%s R (%s)", getRversion(), paste(getRversion(), R.version$platform, R.version$arch, R.version$os))); \
  codename <- system(". /etc/os-release; echo $VERSION_CODENAME", intern = TRUE); \
  options(repos = c(CRAN = sprintf("https://packagemanager.posit.co/cran/__linux__/%s/latest", codename))); \
  install.packages(c("shiny", "dplyr", "stringr", "readr", "httr2", "jsonlite", "magrittr", "openxlsx"), Ncpus = max(1, parallel::detectCores())); \
  if (!all(c("shiny", "dplyr", "stringr", "readr", "httr2", "jsonlite", "magrittr", "openxlsx") %in% rownames(installed.packages()))) quit(status = 1)'

COPY web-schedule/app.R /app/app.R

EXPOSE 3838
CMD ["R", "-e", "shiny::runApp('/app', host = '0.0.0.0', port = 3838)"]
