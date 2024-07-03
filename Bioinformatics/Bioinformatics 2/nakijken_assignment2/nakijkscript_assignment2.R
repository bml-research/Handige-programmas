# script om Opdracht 1 Bioinformatica 2 na te kijken


library(rstudioapi)
library(openxlsx)
library(tibble)

save_path <- setwd(dirname(rstudioapi::getActiveDocumentContext()$path))

# read student answers into df
student_answers <- function(file_to_read) {
  file_to_read <- readWorkbook(paste0(save_path, "/answers_students_assignment2_inhaal.xlsx"), 
                               sheet = "Form1")
  students_df <- file_to_read[, c(2,3,5:ncol(file_to_read))]
  students_df[,5:ncol(students_df)] <- apply(students_df[,5:ncol(students_df)], 2, toupper)
  students_df$Tijd.van.laatste.wijziging <- NULL
  students_df$Begintijd<- strftime(convertToDateTime(students_df$Begintijd), 
                                     format = "%d/%m/%Y %H:%M:%S")
  students_df$Tijd.van.voltooien <- 
    strftime(convertToDateTime(students_df$Tijd.van.voltooien), 
                                    format = "%d/%m/%Y %H:%M:%S")
  return(students_df)
}

input_students <- student_answers()

answers_datasets <- function(input_students) {
  # read answer file and take only datasets used by students
  all_answers <- data.frame(readWorkbook(paste0(save_path, "/all_answers.xlsx"), 1))
  all_answers[,] <- apply(all_answers[,], 2, toupper)
  answer_list <- list()
  for(row in 1:nrow(input_students)) {
    test1 <- all_answers[input_students[row,ncol(input_students)], ]
    answer_list[[row]] <- data.frame(test1)
    
  }
  
  new_df2 <- Reduce(rbind, answer_list)
  return(new_df2)
}

answers <- answers_datasets(input_students)

#check if rows student and answers are the same
nakijken <- function(input_students, answers) {
  output_list <- list()
  for(rijtje in 1:nrow(answers)) {
    rij_antwoord <- answers[rijtje, 2:ncol(answers)]
    rij_student <- input_students[rijtje, 4:(ncol(input_students) - 3)]
    check_antw <- rij_antwoord == rij_student
    new_list <- data.frame()
    weging <- rep(c(1), times = length(rij_antwoord))
    for(t in 1:length(check_antw)) {
      ifelse(check_antw[t] == TRUE, new_list <- append(new_list, c(rij_antwoord[, t], 
                                                                   rij_student[, t], weging[t])), 
             new_list <- append(new_list, c(rij_antwoord[, t], rij_student[, t], 0)))
  
    }
    
    # add score (percentage)
    vragen <- length(new_list) / 3
    score <- sum(as.numeric(new_list[seq(0, length(new_list), 3)]))
    percentage <- round(sum(as.numeric(new_list[seq(0, length(new_list), 
                                                    3)])) / sum(weging) * 100, 0)
    new_list <- append(new_list, c(vragen, score, percentage), after = 0)
  output_list[[rijtje]] <- new_list
  
  }
  
  #create the dataframe with all answers, student answers and points scored
  new_df3 <- Reduce(rbind, output_list)
  new_df4 <- cbind(input_students[, c(15, 3, 16, 17, 1, 2)], new_df3)

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
  
  colnames(new_df4) <- c("Studentnummer", "Naam", "Class", "Dataset", "Starttijd", 
                         "Eindtijd", "Aantal vragen", "Score", "Percentage", headers)
  
  # change columns with points to numeric
  col_sel <- which(startsWith(colnames(new_df4), "Punt") == T)
  new_df4[, c(7:9, col_sel)] <- sapply(new_df4[, c(7:9, col_sel)], as.numeric)
  
  return(new_df4)

}


nagekeken <- nakijken(input_students, answers)

# write results to Excel file
write_file <- function(nagekeken) {
  file_name <- paste0(save_path, "/final_nakijk_assignment2.xlsx")
  new_wb <- createWorkbook(file_name)
  addWorksheet(new_wb, "nakijk_eindopdracht")
  writeData(new_wb, 1, nagekeken)
  
  negStyle <- createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
  posStyle <- createStyle(fontColour = "#006100", bgFill = "#C6EFCE")
  repStyle <- createStyle(fontColour = "#9C6500", bgFill = "#FFEB9C")
  conditionalFormatting(new_wb, 1, 
                        cols = which(names(nagekeken) == "Question1"):ncol(nagekeken),
                        rows = 1:nrow(nagekeken) + 1, type = "between", rule = c(0,10),
                        style = posStyle)
  
  conditionalFormatting(new_wb, 1, 
                        cols = which(names(nagekeken) == "Question1"):ncol(nagekeken),
                        rows = 1:nrow(nagekeken) + 1, rule = "==0",
                        style = negStyle)
  
  conditionalFormatting(new_wb, 1, cols = which(names(nagekeken) == "Percentage"),
                        rows = 1:nrow(nagekeken) + 1, rule = "<50",
                        style = negStyle)
  
  conditionalFormatting(new_wb, 1, cols = which(names(nagekeken) == "Percentage"),
                        rows = 1:nrow(nagekeken) + 1, rule = ">70",
                        style = posStyle)
  
  conditionalFormatting(new_wb, 1, cols = which(names(nagekeken) == "Percentage"),
                        rows = 1:nrow(nagekeken) + 1, type = "between", rule = c(50, 70),
                        style = repStyle)
  
  saveWorkbook(new_wb, file_name, overwrite = T)
  
}


write_file(nagekeken)

