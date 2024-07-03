# script om Enzymkinetiek Opdracht na te kijken



library(openxlsx)
library(rstudioapi)

# location for reading and writing data
work_path <- setwd(dirname(rstudioapi::getActiveDocumentContext()$path))
time_stamp <- format(Sys.Date(), "%Y-%m-%d")
answers_path <- paste0(work_path, "/Nakijken/Answers/")
students_path <- paste0(work_path, "/Nakijken/Files/")
save_path <- paste0(work_path, "/Nakijken/")

# make a dataframe from the answers of the students
## first column contains the number of the dataset used

student_answers <- function(files_to_read) {
  files_to_read <- list.files(path = students_path, pattern = ".xlsx")
  answer_students <- list()
  for(f in files_to_read) {
    data_students <- readWorkbook(paste0(students_path, f), 3, rows = c(21,22))
    opdracht1 <- readWorkbook(paste0(students_path,f), 2, rows = c(14,15))
    opdracht2 <- data.frame(lapply(data_students, 
                                     function(x) if(is.numeric(x)) round(x, 1) else x))
    opdracht3 <- readWorkbook(paste0(students_path, f), 3, rows = c(27,28))
    sheet_name <- getSheetNames(paste0(students_path, f))
    student_nr <- as.numeric(gsub("\\D","", f))
    data_name <- as.numeric(gsub("\\D", "", sheet_name[1]))
    data_students <- cbind(student_nr, data_name, opdracht1, opdracht2, opdracht3)
    answer_students[[f]] <- data.frame(data_students)

  }
  
  students_df <- Reduce(rbind, answer_students)
  return(students_df)
}

input_student <- student_answers()


# antwoorden voor het nakijken
## alleen antwoorden voor datasets die door studenten zijn gemaakt
answers_datasets <- function(input_student) {
  # read answer file
  all_answers_comp <- readWorkbook(paste0(answers_path, "all_answers_comp.xlsx"), 1)
  all_answers_noncomp <- readWorkbook(paste0(answers_path, "all_answers_noncomp.xlsx"), 1)
  all_answers <- rbind(all_answers_comp, all_answers_noncomp)
  #testje <- as.numeric(unlist(all_answers))
  #answers_2_dec <- round(all_answers, 2)
  all_answers <- data.frame(lapply(all_answers, 
                                   function(x) if(is.numeric(x)) round(x, 1) else x))
  #extract rows with answers that are made by students and make a data frame 
  answer_list <- list()
  for(row in 1:nrow(input_student)) {
    test1 <- all_answers[input_student[row,2], ]
    answer_list[[row]] <- data.frame(test1)
    
  }
  
  new_df2 <- Reduce(rbind, answer_list)
  return(new_df2)
}

answers <- answers_datasets(input_student)
#check if rows student and answers are the same

nakijken <- function(input_student, answers) {
  output_list <- list()
  for(rijtje in 1:nrow(answers)) {
    rij_antwoord <- answers[rijtje, 2:6]
    rij_student <- input_student[rijtje, 7:11]
    opdr1 <- input_student[rijtje, 3:6]
    opdr2 <- input_student[rijtje, 7:10]
    opdr3 <- input_student[rijtje, 11]
    marge_opdr1 <- list()
    marge_opdr2 <- list()
    for(x in opdr1) {
      marge_opdr1 <- append(marge_opdr1, data.table::between(x, 0.8*rij_antwoord[match(x, opdr1)],
                                   1.2*rij_antwoord[match(x, opdr1)]))
    }
    
    for(y in opdr2) {
      marge_opdr2 <- append(marge_opdr2, data.table::between(y, 0.95*rij_antwoord[match(y, opdr2)],
                                                             1.05*rij_antwoord[match(y, opdr2)]))
    }

    check_opdr3 <- rij_antwoord[, 5] == opdr3
    new_list <- data.frame()
    #weging <- c(1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1)
    
    for(u in 1:length(marge_opdr1)) {
      ifelse(marge_opdr1[u] == T, new_list <- append(new_list, c(rij_antwoord[, u],
                                                       opdr1[, u], 0.25)),
             new_list <- append(new_list, c(rij_antwoord[, u], opdr1[, u], 0)))
    }

    for(p in 1:length(marge_opdr2)) {
      ifelse(marge_opdr2[p] == T, new_list <- append(new_list, c(rij_antwoord[, p],
                                                                 opdr2[, p], 1)),
             new_list <- append(new_list, c(rij_antwoord[, p], opdr2[, p], 0)))
    }
    
        
    for(t in 1:length(check_opdr3)) {
      ifelse(check_opdr3[t] == TRUE, new_list <- append(new_list, c(rij_antwoord[, 5], 
                                                                   opdr3[t], 2)), 
             new_list <- append(new_list, c(rij_antwoord[, 5], opdr3[t], 0)))
    }

    score <- sum(na.omit(as.numeric(new_list[seq(0, length(new_list), 3)])))
    new_list <- append(new_list, score)
    
    output_list[[rijtje]] <- new_list
  }
  
  new_df3 <- Reduce(rbind, output_list)
  new_df4 <- cbind(input_student[, 1:2], new_df3)
  colnames(new_df4) <- c("Studentnummer", "Dataset", "1. Vmax_zonder_answer",
                         "1. Vmax_zonder_student", "1. punt","1. Vmax_met_answer", 
                         "1. Vmax_met_student", "1. punt", "1. Km_zonder_answer", 
                         "1. Km_zonder_student", "1. punt", "1. Km_met_answer", 
                         "1. Km_met_student", "1. punt", "2. Vmax_zonder_answer",
                         "2. Vmax_zonder_student", "2. punt","2. Vmax_met_answer", 
                         "2. Vmax_met_student", "2. punt", "2. Km_zonder_answer", 
                         "2. Km_zonder_student", "2. punt", "2. Km_met_answer", 
                         "2. Km_met_student", "2. punt", "3. inhibitor_answer",
                         "3. inhibitor_student", "3. punt", "totaal punten")
  
  return(as.data.frame(new_df4))

}

nagekeken <- nakijken(input_student, answers)


# wegschrijven van nagekeken data
write.xlsx(nagekeken, file = paste0(save_path, "nagekeken_weighted_", 
                                    time_stamp, ".xlsx"))
