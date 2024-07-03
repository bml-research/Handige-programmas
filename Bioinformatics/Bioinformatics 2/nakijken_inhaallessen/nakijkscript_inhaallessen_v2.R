# script om inhaallessen Bioinformatica 2 na te kijken



library(openxlsx)
library(tibble)
library(rstudioapi)

save_path <- setwd(dirname(rstudioapi::getActiveDocumentContext()$path))

student_answers <- function(file_to_read) {
  file_to_read <- readWorkbook(paste0(save_path, "/Lesson 1-14_Answer sheet_inhaal.xlsx"), sheet = "Form1")
  students_df <- file_to_read[, c(2,3,5:(ncol(file_to_read)))]
  students_df[, 1] <- strftime(convertToDateTime(students_df[, 1]), 
                                     format = "%d/%m/%Y %H:%M:%S")
  students_df[, 2] <- strftime(convertToDateTime(students_df[, 2]), 
             format = "%d/%m/%Y %H:%M:%S")

  return(students_df)  
}

input_students <- student_answers()

answers_datasets <- function(input_students) {
  # read answer file and take only datasets used by students
  all_answers <- readWorkbook(paste0(save_path, "/inhaal_all_answers.xlsx"), 1)
  answer_list <- list()
  for(row in 1:nrow(input_students)) {
    test1 <- all_answers[input_students[row,ncol(input_students)], ]
    answer_list[[row]] <- data.frame(test1)
    
  }
  
  new_df2 <- Reduce(rbind, answer_list)
  return(new_df2)
}


answers <- answers_datasets(input_students)


nakijken <- function(input_students, answers) {
  output_list <- list()
  for(rijtje in 1:nrow(answers)) {
    rij_antwoord <- answers[rijtje, 2:ncol(answers)]
    rij_student <- input_students[rijtje, 4:(ncol(input_students) - 5)]
    check_antw <- rij_antwoord == rij_student
    new_list <- data.frame()
    #weging <- c(1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1)
    for(t in 1:length(check_antw)) {
      ifelse(check_antw[t] == TRUE, new_list <- append(new_list, c(rij_antwoord[, t], 
                                                                   rij_student[, t], 1)), 
             new_list <- append(new_list, c(rij_antwoord[, t], rij_student[, t], 0)))
      ifelse(is.na(check_antw[t]), new_list <- append(new_list, c(rij_antwoord[, t],
                                                                  rij_student[, t], NA)),
             NA)
    }
    
    # add score (percentage)
    vragen <- sum(is.na(check_antw) == FALSE)
    score <- sum(na.omit(as.numeric(new_list[seq(0, length(new_list), 3)])))
    percentage <- round(score / vragen * 100, 2)
    new_list <- append(new_list, c(vragen, score, percentage), after = 0)
    
    output_list[[rijtje]] <- new_list
    
  }
  
  #create the dataframe with all answers, student answers and points scored
  new_df3 <- Reduce(rbind, output_list)
  get_index <- c()
  sel_col <- c("Student.Number", "Name", "Class", "Group", "Group2", "Lesson", "Start.time", "Completion.time")
  for(z in sel_col) {
    get_index <- append(get_index, which(colnames(input_students) == z))
  }
  new_df4 <- cbind(input_students[, get_index], new_df3)
  new_df5 <- is.na(new_df4$Group)
  new_df4$Group[new_df5] <- new_df4$Group2[new_df5]
  new_df4$Group2 <- NULL
  
  aantal_vragen <- c()
  aantal_antwoorden <- c()
  aantal_punten <- c()
  headers <- c()
  for(y in 1:ncol(rij_antwoord)) {
    aantal_vragen <- append(aantal_vragen, paste0("Question", y))
    aantal_antwoorden <- append(aantal_antwoorden, paste0("Answer", y))
    aantal_punten <- append(aantal_punten, paste0("Punt_vraag", y))
    my_list <- list(aantal_vragen, aantal_antwoorden, aantal_punten)
  }  
  for(j in 1:length(aantal_vragen)) {
    headers <- append(headers, sapply(my_list, "[[", j))
    
  }
  
  colnames(new_df4) <- c("Studentnummer", "Naam", "Class", "Group", "Les", "Starttijd", 
                         "Eindtijd", "Aantal vragen", "Score", "Percentage", headers)
  
  # change columns with points to numeric
  col_sel <- which(startsWith(colnames(new_df4), "Punt") == T)
  new_df4[is.na(new_df4)] <- ""
  new_df4[, c(8:10, col_sel)] <- sapply(new_df4[, c(8:10, col_sel)], as.numeric)
  new_df4[, col_sel] <- replace(new_df4[, col_sel], is.na(new_df4[, col_sel]), 100000)
  return(new_df4)
  
}


nagekeken <- nakijken(input_students, answers)

# write results to Excel file
write_file <- function(nagekeken) {
  file_name <- paste0(save_path, "/results_inhaallessen_", 
                      format(Sys.Date(), "%Y%m%d"), ".xlsx")
  new_wb <- createWorkbook(file_name)
  addWorksheet(new_wb, "nakijk_inhaallessen")
  writeData(new_wb, 1, nagekeken)
  score <- paste0("=SUMIF(K", seq(2, nrow(nagekeken) + 1, 1), ":ER", 
                  seq(2, nrow(nagekeken) + 1, 1), ",1)")
  writeFormula(new_wb, 1, x = score, startCol = 9, startRow = 2)
  perc <- paste0("=ROUND(I", seq(2, nrow(nagekeken) + 1, 1), "/H", 
                 seq(2 ,nrow(nagekeken)), "*100,2)")
  writeFormula(new_wb, 1, x = perc, startCol = 10, startRow = 2)
  addFilter(new_wb, 1, rows = 1, cols = 1:ncol(nagekeken))
  
  negStyle <- createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
  posStyle <- createStyle(fontColour = "#006100", bgFill = "#C6EFCE")
  repStyle <- createStyle(fontColour = "#9C6500", bgFill = "#FFEB9C")
  noStyle <- createStyle(fontColour = "#FFFFFF")
  conditionalFormatting(new_wb, 1, cols = which(colnames(nagekeken) == "Punt_vraag1"):ncol(nagekeken),
                        rows = 1:nrow(nagekeken) + 1, type = "between", rule = c(0,10),
                        style = posStyle)
  
  conditionalFormatting(new_wb, 1, cols = which(colnames(nagekeken) == "Punt_vraag1"):ncol(nagekeken),
                        rows = 1:nrow(nagekeken) + 1, rule = '==100000',
                        style = noStyle)
  
  conditionalFormatting(new_wb, 1, cols = which(colnames(nagekeken) == "Punt_vraag1"):ncol(nagekeken),
                        rows = 1:nrow(nagekeken) + 1, rule = "==0",
                        style = negStyle)

  conditionalFormatting(new_wb, 1, cols = which(colnames(nagekeken) == "Percentage"),
                        rows = 1:nrow(nagekeken) + 1, rule = "<50",
                        style = negStyle)
  
  conditionalFormatting(new_wb, 1, cols = which(colnames(nagekeken) == "Percentage"),
                        rows = 1:nrow(nagekeken) + 1, rule = ">70",
                        style = posStyle)
  
  conditionalFormatting(new_wb, 1, cols = which(colnames(nagekeken) == "Percentage"),
                        rows = 1:nrow(nagekeken) + 1, type = "between", rule = c(50, 70),
                        style = repStyle)
  
  saveWorkbook(new_wb, file_name, overwrite = T)
  
}


write_file(nagekeken)



