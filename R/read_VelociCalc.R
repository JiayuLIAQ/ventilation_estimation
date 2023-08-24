source("R/functions.R")

path_xls <- data.table(path = list.files(no.. = FALSE, full.names = TRUE, recursive = TRUE,
                                         pattern = "\\.xls$")) 

path_xls[path %like% "raw_data", data_type := "raw_data"]
path_xls[path %like% "clean_data", data_type := "clean_data"]

path_xls[, scenario := str_match(path, paste0(data_type,"/(.*?)/") )[,2] ]
path_xls[, equip := str_match(path, paste0(scenario,"/(.*?)/") )[,2] ]
# path_xls[, folder := str_match(path, paste0(equip,"/(.*?)/") )[,2] ]

read_velocicalc_with_co2_2 <- function(path){
  
  dt <- fread(path, sep = "")
  
  dt[, n_char := str_count(`LogDat2 Data File`)]
  dt[, n_tab := str_count(`LogDat2 Data File`, "\t")]
  dt[, n_digit := str_count(`LogDat2 Data File`, "[^0-9]+")]
  dt <- dt[n_char %in% c(34:52) & n_digit %in% c(10:17) & n_tab %in% c(4:8)]
  if (unique(dt$n_tab) == 5) {
    dt[, c('date', 'time', 'co2', 'temp', 'rh', 'co') := tstrsplit(`LogDat2 Data File`, "\t")]
    dt[, c('co2', 'temp', 'rh', 'co') :=lapply(.SD, as.numeric), 
       .SDcols=c('co2', 'temp', 'rh', 'co')]
  } else if (unique(dt$n_tab) == 6) {
    dt[, c('date', 'time', 'co2', 'p_inH20', 'temp', 'rh', 'co') := tstrsplit(`LogDat2 Data File`, "\t")]
    dt[, c('co2', 'p_inH20', 'temp', 'rh', 'co') :=lapply(.SD, as.numeric), 
       .SDcols=c('co2', 'p_inH20', 'temp', 'rh', 'co')]
  } else if (unique(dt$n_tab) == 7) {
    dt[, c('date', 'time', 'co2', 'p_inH20', 'temp', 'rh', 'co', 'p_hPa') := tstrsplit(`LogDat2 Data File`, "\t")]
    dt[, c('co2', 'p_inH20', 'temp', 'rh', 'co', 'p_hPa') :=lapply(.SD, as.numeric), 
       .SDcols=c('co2', 'p_inH20', 'temp', 'rh', 'co', 'p_hPa')]
  }
  dt[, datetime := dmy_hms(paste(date, time))]
  dt[, c('LogDat2 Data File', 'n_char', 'n_tab', 'n_digit', 'date', 'time') := NULL]
  
  dt[, equip_unit := file_path_sans_ext(basename(path)) %>% str_sub(start = 4L, end = 4L)]
}

dt_all <- map(path_xls$path, read_velocicalc_with_co2_2) %>% rbindlist(fill = T) %>% .[!is.na(datetime)]

dt_all[, c("p_inH20","p_hPa") := NULL ]


dt_all[date(datetime) == "2023-08-22" & equip_unit != "U"] %>% ggplot(aes(datetime, co2, color = equip_unit)) + geom_point() +
  ylab(expression(bold("C"*O[2]*" concentration (ppm)"))) +
  xlab("Time of the day") +
  LJYtheme_basic
