library(xlsx)
library(openxlsx)
library(gtools)
library(stringr)

setwd("/Users/markuiz/PycharmProjects/Eigen projectjes/Nakijken_Bioinformatica/Eindbeoordeling/input")
read_path <- "/Users/markuiz/PycharmProjects/Eigen projectjes/Nakijken_Bioinformatica/Eindbeoordeling/input/"
save_path <- "/Users/markuiz/PycharmProjects/Eigen projectjes/Nakijken_Bioinformatica/Eindbeoordeling/output/"

files_to_read <- mixedsort(list.files(
  path = read_path, 
  pattern = ".xlsx"))
df_percentages <- list()
for(f in files_to_read) {
  lespercentages <- read.xlsx(f, 1, cols = c(1:4, 9))
  df_percentages[[f]] <- data.frame(lespercentages)
  #df_percentages[[f]]$Percentage <- as.numeric(str_extract(f, "\\d+"))
  #colnames(df_punten[[f]]) <- c("Studentnummer", "Naam", "Percentage", "Les")
}

merged_df <- Reduce(function(...) merge(..., by = "Naam", all = T), 
                    df_percentages)

# studenten die eerste les missen worden anders niet meegenomen met de merge
merged_df[2:3][is.na(merged_df[2:3])] <- 0

df_new <- merged_df[, c(1:4, seq(5, ncol(merged_df), 4))]
df_new[, 5:17] <- round(df_new[, 5:17], 2)
df_new <- subset(df_new, df_new$Naam != "Mark Sibbald")
df_new <- subset(df_new, df_new$Klas.x != "BOVD3")
colnames(df_new) <- c("Naam", "Studentnummer", "Class", "Group", paste0("Les ", 
                                                      rep(seq_along(files_to_read))))

df_new$mean <- round(rowMeans(df_new[5:length(df_new)]), 1)
df_new$Beoordeling <- ifelse(apply(df_new[5:17], 1, function(x) any(x < 70)) == T,
                             "Onvoldoende", "Voldoende")
df_new$Beoordeling[is.na(df_new$mean)] <- "Onvoltooid"

# write results to Excel file
write_file <- function(df_new) {
  file_name <- paste0(save_path, "eindbeoordeling.xlsx")
  new_wb <- createWorkbook(file_name)
  addWorksheet(new_wb, "eindbeoordeling")
  writeData(new_wb, 1, df_new)
  
  negStyle <- createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
  posStyle <- createStyle(fontColour = "#006100", bgFill = "#C6EFCE")
  repStyle <- createStyle(fontColour = "#9C6500", bgFill = "#FFEB9C")
  conditionalFormatting(new_wb, 1, cols = 5:(ncol(df_new) - 1),
                        rows = 1:nrow(df_new) + 1, type = "between", rule = c(0,49),
                        style = negStyle)
  
  conditionalFormatting(new_wb, 1, cols = 5:(ncol(df_new) - 1),
                        rows = 1:nrow(df_new) + 1, type = "between", rule = c(50,69.99),
                        style = repStyle)
  
  conditionalFormatting(new_wb, 1, cols = 5:(ncol(df_new) - 1),
                        rows = 1:nrow(df_new) + 1, type = "between", rule = c(70,100),
                        style = posStyle)
  
  conditionalFormatting(new_wb, 1, cols = ncol(df_new),
                        rows = 1:nrow(df_new) + 1, rule = "Voldoende",
                        style = posStyle)
  
  conditionalFormatting(new_wb, 1, cols = ncol(df_new),
                        rows = 1:nrow(df_new) + 1, rule = "Onvoldoende",
                        style = negStyle)

  conditionalFormatting(new_wb, 1, cols = ncol(df_new),
                        rows = 1:nrow(df_new) + 1, rule = "Onvoltooid",
                        style = repStyle)
  
  saveWorkbook(new_wb, file_name, overwrite = T)
  
}

write_file(df_new)
