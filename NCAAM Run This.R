run_libs = function(){
  
  library(tidyverse)
  library(ggpubr)
  library(readxl)
  library(writexl)
  library(xlsx)
  library(reshape2)
  library(openxlsx)
}
run_libs()


getwd()
setwd("D:/Projects/Sports Data")



### --- RUN THIS --- ###
run_lines = function(sheet_name){
  
  lines = read.xlsx("NCAAM.xlsx", sheet = sheet_name)
  rankings = read.xlsx("NCAAM.xlsx", sheet = "Rankings")
  mapping = read.xlsx("NCAAM.xlsx", sheet = "Mapping")
  rankings["Points.Rank"] = rankings$Rk / 10
  
  lines$Home.Away = "Home"
  for (i in seq(1,nrow(lines), 2)){
    lines$Home.Away[i] = "Away"
  }
  
  if(sheet_name == "Lines1"){
    lines$Tomorrow = str_replace(lines$Tomorrow, "_x000D_\n", " ")
    lines$Spread = str_replace_all(lines$Spread, "_x000D_\n", " ")
    lines$Total = str_replace_all(lines$Total, "_x000D_\n", " ")
    
    for (i in 1:nrow(lines)){
      if(i %% 2 == 1){
        lines$Opponent[i] = lines$Tomorrow[i+1]
      }
      else{
        lines$Opponent[i] = lines$Tomorrow[i-1]
      }
    }
    
    lines
    
    lines$Time = substr(lines$Tomorrow, 0, 7)
    lines$Time = trimws(lines$Time, which = c("both"))
    lines$Time = format(as.POSIXct(lines$Time, format = "%I:%M%p", tz = "UTC"), tz = "America/New_York", format = "%H:%M:%S")
    
    # lines$Tomorrow = str_replace(lines$Tomorrow, "\r\n", " ")
    # lines$Spread = str_replace_all(lines$Spread, "\r\n", " ")
    # lines$Total = str_replace_all(lines$Total, "\r\n", " ")
    # lines
    
    lines = separate(lines, Spread, into = c("Spread", "Spread Odds"), sep = " ")
    lines$Spread = str_replace(lines$Spread, "pk", "0")
    lines_final = separate(lines, "Total", into = c("O/U Amount", "O/U Odds"), sep = " ")
    
    lines_final2 = lines_final
    schedule = merge(lines_final, lines_final2, by.x = "Tomorrow", by.y = "Opponent")
    unique_sched = schedule %>% filter (Home.Away.x == "Away")
    unique_sched = separate(unique_sched, Tomorrow, into = c("Time", "Away"), sep = "[0-9][AP]M ")
    clean_sched = separate(unique_sched, Opponent, into = c("Time.y", "Home"), sep = "[0-9][AP]M ")
  }
  if(sheet_name == "Lines0"){
    lines$Today = str_replace(lines$Today, "_x000D_\n", " ")
    lines$Spread = str_replace_all(lines$Spread, "_x000D_\n", " ")
    lines$Total = str_replace_all(lines$Total, "_x000D_\n", " ")
    
    for (i in 1:nrow(lines)){
      if(i %% 2 == 1){
        lines$Opponent[i] = lines$Today[i+1]
      }
      else{
        lines$Opponent[i] = lines$Today[i-1]
      }
    }
    
    lines
    
    lines$Time = substr(lines$Today, 0, 7)
    lines$Time = trimws(lines$Time, which = c("both"))
    lines$Time = format(as.POSIXct(lines$Time, format = "%I:%M%p", tz = "UTC"), tz = "America/New_York", format = "%H:%M:%S")
    
    # lines$Today = str_replace(lines$Today, "\r\n", " ")
    # lines$Spread = str_replace_all(lines$Spread, "\r\n", " ")
    # lines$Total = str_replace_all(lines$Total, "\r\n", " ")
    # lines
    
    lines = separate(lines, Spread, into = c("Spread", "Spread Odds"), sep = " ")
    lines$Spread = str_replace(lines$Spread, "pk", "0")
    lines_final = separate(lines, "Total", into = c("O/U Amount", "O/U Odds"), sep = " ")
    
    lines_final2 = lines_final
    schedule = merge(lines_final, lines_final2, by.x = "Today", by.y = "Opponent")
    unique_sched = schedule %>% filter (Home.Away.x == "Away")
    unique_sched = separate(unique_sched, Today, into = c("Time", "Away"), sep = "[0-9][AP]M ")
    clean_sched = separate(unique_sched, Opponent, into = c("Time.y", "Home"), sep = "[0-9][AP]M ")
    
    
  }
  
  clean_sched = merge(clean_sched, mapping, by.x = "Away", by.y = "DraftKings", all.x = TRUE)
  print(clean_sched[is.na(clean_sched$Ken.Pom), ])
  
  clean_sched = merge(clean_sched, mapping, by.x = "Home", by.y = "DraftKings", all.x = TRUE)
  print(clean_sched[is.na(clean_sched$Ken.Pom.y), ])
  
  clean_sched
  m1 = merge(clean_sched, rankings, by.x = "Ken.Pom.x", by.y = "Team", all.x = TRUE)
  
  m2 = merge(m1, rankings, by.x = "Ken.Pom.y", by.y = "Team", all.x = TRUE)
  m2[is.na(m2$Strength.of.Schedule.AdjEM.y), ]
  
  m2$Rank.Difference = m2$Points.Rank.x - m2$Points.Rank.y
  m2$Home.Buffer = 5
  m2$New.Spread = m2$Rank.Difference + m2$Home.Buffer
  m2$Spread.Difference = m2$New.Spread - as.numeric(m2$Spread.x)
  
  m2$Flag = abs(m2$Spread.Difference) > 3
  
  m2$Lean = ifelse(m2$Spread.Difference > 3, "Home", ifelse(m2$Spread.Difference < -3, "Away", "N/A"))
  
  
  m2$Away.Projection = round(m2$AdjD.y/m2$AdjO.x*(m2$AdjO.x*m2$AdjT.x/100), 1)
  m2$Home.Projection = round(m2$AdjD.x/m2$AdjO.y*(m2$AdjO.y*m2$AdjT.y/100), 1)
  m2$Total.Projection = m2$Home.Projection + m2$Away.Projection
  m2$Over.Under = as.numeric(substring(m2$`O/U Amount.x`, 3))
  m2$Over.Under.Difference = m2$Over.Under - m2$Total.Projection
  m2$Projection.Difference = m2$Away.Projection - m2$Home.Projection

  
  m2$OU.Lean = ifelse(m2$Over.Under.Difference < -10, "Over", ifelse(m2$Over.Under.Difference > 10, "Under", "N/A")) 
  
  colnames(m2)
  
  filter_cols = c( "Time.x", "Home", "Away", "Lean", "Points.Rank.y", "Points.Rank.x", 
                   "Rank.Difference", "Home.Buffer", "New.Spread", "Spread.Difference",
                   "Spread.y", "Spread.x", "Ken.Pom.y", "Ken.Pom.x", "Rk.y", "Rk.x", "Over.Under", "O/U Odds.x",
                   "AdjO.x", "AdjD.y", "AdjO.y", "AdjD.x", "AdjT.y", "AdjT.x", 
                   "Home.Projection", "Away.Projection", "Total.Projection", 
                   "OU.Lean", "Over.Under.Difference", "Projection.Difference"
  )
  
  
  results = m2[, filter_cols]
  
  colnames(results) = c( "Time.EST", "Home", "Away",  "Lean", "Points.Rank.Home", "Points.Rank.Away", 
                         "Rank.Difference", "Home.Buffer", "New.Spread", "Spread.Difference",
                         "Spread.Home", "Spread.Away", "Ken.Pom.Home", "Ken.Pom.Away","Rk.Home", "Rk.Away", "O/U Line", "O/U Odds",
                         "AdjO.Away", "AdjD.Home", "AdjO.Home", "AdjD.Away", "AdjT.Home", "AdjT.Away",
                         "Home.Projection", "Away.Projection", "Total.Projection", "O/U Lean", "O/U Difference", "Projection.Difference")
  
  results = results[, c( "Time.EST", "Away", "Home",  "Lean", "New.Spread", "Spread.Difference", "Spread.Away", "Spread.Home",
                         "Projection.Difference", "Away.Projection", "Home.Projection", "Total.Projection", 
                         "O/U Line", "O/U Difference", "O/U Lean",
                         "Points.Rank.Away", "Points.Rank.Home", "Rank.Difference", "Home.Buffer", 
                         "Ken.Pom.Away", "Ken.Pom.Home","Rk.Away", "Rk.Home",
                         "AdjO.Away", "AdjD.Home", "AdjO.Home", "AdjD.Away", "AdjT.Away", "AdjT.Home"
  )]
  return(results)
}

Lines1 = run_lines("Lines1")
Lines0 = run_lines("Lines0")
final_results = rbind(Lines0, Lines1)
write_xlsx(final_results, paste(format(Sys.time(), "%Y_%m_%d"), "NCAAM Basketball" , format(Sys.time(), "%H_%M_%S"), ".xlsx"))
