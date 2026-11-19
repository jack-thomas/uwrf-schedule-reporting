FROM rocker/shiny-verse:4.4.1

RUN R -q -e 'install.packages(c("httr2", "openxlsx"), repos = "https://cloud.r-project.org")'

COPY web-schedule/app.R /srv/shinyapps/web-schedule/app.R

EXPOSE 3838

CMD ["R", "-e", "shiny::runApp('/srv/shinyapps/web-schedule', host = '0.0.0.0', port = 3838)"]
