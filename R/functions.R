Sys.setenv(LANG = "en")
library(RODBC)

library(readxl)
library(magrittr)
library(data.table)
library(lubridate)
library(ggplot2)
library(patchwork)
library(xlsx)
library(tools)
library(purrr)

library(janitor)

library(plotly)
library(ggpmisc)

library(stringr)


# VelociCalc with co2----------------------------------------------

read_velocicalc_with_co2 <- function(path){
  
  dt <- fread(path, skip = 29) [-1]
  
  dt[, datetime := dmy_hms(paste(Date, Time))]
  
  dt[, c("CO2","T","H", "CO") :=lapply(.SD, as.numeric), .SDcols=c("CO2","T","H","CO")]
  
  dt %>% setnames(old = c("CO2","T","H", "CO"), new = c("co2", "temperature", "RH","co"))
  
  dt[, c("Date","Time") := NULL] 
  
  dt[, equip_unit := file_path_sans_ext(basename(path))]
  
}

# VelociCalc with no co2----------------------------------------------

read_velocicalc <- function(path){
  
  dt <- fread(path, skip = 33) [-1]
  
  dt[, datetime := dmy_hms(paste(Date, Time))]
  
  dt[, c("Vel","T","H") :=lapply(.SD, as.numeric), .SDcols=c("Vel","T","H")]
  
  dt %>% setnames(old = c("Vel","T","H"), new = c("vel", "temperature", "RH"))
  
  char_cols_T_F <- sapply(dt, is.character)
  names_char <- dt %>% names %>% .[char_cols_T_F]
  # if(length > 1)
  dt[, (names_char) := NULL] 
  
  dt[, equip_unit := file_path_sans_ext(basename(path))]
  
}

# innova----------------------------------------------
read_mdb_and_write_xlsx <- function(path){
  
  conn <- odbcDriverConnect(paste0("Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=",path))
  
  dt_tb_list <- RODBC::sqlTables(conn) %>% setDT
  
  dt_tb_list[TABLE_NAME %like% "airr"]$TABLE_NAME
  
  read_one_table <- function(table_name_) {
    data_ <- RODBC::sqlFetch(conn, table_name_) %>% setDT
    data_[, table_name := table_name_]
  }
  
  test <- map(dt_tb_list[TABLE_NAME %like% "airr"]$TABLE_NAME, 
              read_one_table ) %>% rbindlist %>% .[, edit := NULL]
  
  n_tables <- unique(test$table_name) %>% length 
  table_names <- unique(test$table_name)
  
  # class(test)
  for(i in 1:n_tables){
    xlsx::write.xlsx(test[table_name == table_names[i]], 
                     file=paste0(file_path_sans_ext(path),".xlsx"), 
                     sheetName=table_names[i], append=TRUE, row.names=FALSE)
  }
  
}

read_mdb <- function(path){
  
  conn <- odbcDriverConnect(paste0("Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=",path))
  
  dt_tb_list <- RODBC::sqlTables(conn) %>% setDT
  
  dt_tb_list[TABLE_NAME %like% "airr"]$TABLE_NAME
  
  read_one_table <- function(table_name_) {
    data_ <- RODBC::sqlFetch(conn, table_name_) %>% setDT
    data_[, table_name := table_name_]
  }
  
  dt <- map(dt_tb_list[TABLE_NAME %like% "airr"]$TABLE_NAME, 
            read_one_table ) %>% rbindlist 
  dt
}


fmt <- paste0("%.", 2, "f")

LJYtheme_basic <- theme(
  plot.title = element_text(size = 12, vjust = 0, face = "bold"),
  axis.text.x = element_text(size = 12, hjust=.5, vjust=1, colour="black"),
  axis.text.y = element_text(size = 12, hjust=1, vjust=.5, colour="black"),
  axis.title.y = element_text(size = 12, color = "black", face = "bold", vjust = 0.5, hjust = 0.5),
  axis.title.x = element_text(size = 12, color = "black", face = "bold", vjust = 0.5, hjust = 0.5),
  axis.line = element_line(color = "black"),
  # panel.grid.major=element_blank(),
  # panel.grid.major = element_line(colour="#f0f0f0"),
  panel.grid.major = element_line(colour="#f7f7f7"),
  # panel.grid.major = element_blank(),
  # panel.grid.minor=element_line(colour="#f0f0f0",size = 0.1),
  panel.grid.minor=element_blank(),
  # panel.background=element_rect(fill='white',colour='black'),
  legend.text = element_text(size = 12),
  legend.key = element_rect(colour = NA, fill = "white"),
  panel.background = element_blank(),
  # legend.position = "bottom",
  # legend.direction = "horizontal",
  # legend.key.size= unit(0.3, "cm"),
  # legend.margin = margin(0,0,0,0,"cm"),
  legend.title = element_text(face = "bold", size = 12),
  strip.background = element_rect(colour= NA, fill="#f0f0f0"),
  strip.text = element_text(face = "bold", size = 12)
)




# conditions-------------------------------------------------
# built_env_variables <- c("CFishmonger",
#                          "widest_aisle",
#                          "freezer_area",
#                          "freezer_num2")
# 
# built_env_labels <- c(expression("Number of C+ fishmongers"),
#                       expression("Widest aisle in market (m)"),
#                       expression("Area ("*m^2*")"),
#                       expression("Number of freezers")
# )
# 
# built_env_names <- set_names(built_env_labels, built_env_variables)

# colors_manual_conditions <- c(   
#   AC_on          ="#1f78b4",
#   AC_off         ="#a6cee3",
#   window_open    ="#fb9a99",
#   window_closed  ="#e31a1c",
#   fan_on         ="#33a02c",
#   fan_off        ="#b2df8a",
#   occupied       ="#ff7f00",
#   unoccupied     ="#fdbf6f",
#   seated         ="#999999",
#   sleep          ="#7570b3",
#   walking        ="#e7298a",
#   making_bed     ="#a6761d",
#   using_bathroom ="#1b9e77",
#   cleaning       ="#ffff33",
#   empty          ="white"
# )
# 
# colors_manual_operation <- c(
#   AC_on          ="#b3cde3",
#   fan_operating  ="#ccebc5",
#   occupied       ="#decbe4",
#   window_open    ="#fed9a6"
# )
# 

