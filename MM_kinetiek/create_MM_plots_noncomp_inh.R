library(xlsx)
library(openxlsx)
library(data.table)
library(reshape2)
library(ggplot2)
library(ggpmisc)
library(rstudioapi)

# location for reading and writing data
work_path <- setwd(dirname(rstudioapi::getActiveDocumentContext()$path))

# create data/dataframes for Lineweaver-Burk, Eadie-Hofstee and Hanes-Woolf plots
# non-competitive inhibition

create_data <- function(km) {
  
  # create answers (and dataframe of answers)
  km <- runif(1, 0.1, 10.0)
  vmax <- km * runif(1, 1.0, 50.0)
  vmax_inh <- vmax * (runif(1, 0.3, 0.9))
  answers <- c(vmax, vmax_inh, km, km)
  answers_matrix <- round(matrix(answers, ncol = 4, byrow = TRUE), 3)
  colnames(answers_matrix) <- c("vmax (mmol/min)", "vmax_inh (mmol/min)", "km (mM)", "km_inh (mM)")
  answers_df <- as.data.frame(answers_matrix, optional = TRUE)
  answers_df$Answer <- "F) Niet-competitieve inhibitor; Vmax verandert, Km blijft gelijk"
  
  # create data from random value of Km
  conc_S <- answers[3] * c(0.5, 1.0, 2.0, 4.0, 6.0, 10.0)
  vel <- c(answers[1] * (conc_S/(answers[3] + conc_S)))
  vel_inh <- c(answers[2] * (conc_S/(answers[4] + conc_S)))
  MM_data <- c(conc_S, vel, vel_inh)
  
  # make matrix and dataframe from one set of data points 
  MM_data_matrix <- round(matrix(MM_data, nrow = 6, byrow = FALSE), 3)
  colnames(MM_data_matrix) <- c("[S] (mM)", "v (mmol/min)", "v_inh (mmol/min)")
  MM_data_df <- as.data.frame(MM_data_matrix, optional = TRUE)
  MM_data_df_melted <- reshape2::melt(MM_data_df, id = "[S] (mM)")
  
  #create LB data
  reciprocal_data <- round(1 / MM_data_df, 3)
  colnames(reciprocal_data) <- c("1/[S] (1/mM)", "1/v (min/mmol)", "1/v_inh (min/mmol)")
  LB_df <- as.data.frame(reciprocal_data, optional = TRUE)
  LB_df_melted <- reshape2::melt(LB_df, id = "1/[S] (1/mM)")
  
  
  # plot v vs [S]
  plot_MM <- ggplot() + geom_point(data = MM_data_df_melted, aes(x = `[S] (mM)`, y = value)) +
    geom_line(data = MM_data_df_melted, aes(x = `[S] (mM)`, y = value, colour = variable)) +
    geom_hline(yintercept = 0.0, colour = "black") +
    geom_vline(xintercept = 0.0, colour = "black") +
    xlim(0, 1.1 * max(conc_S)) + ylim(0, answers[1]) +
    labs(x = "[S] (mM)", y = "v (mmol/min)") +
    ggtitle("Michaelis-Menten Plot") + theme(plot.title = element_text(size = 28,face = "bold"))
  
  
  # plot LB plot
  plot_LB <- ggplot(data = LB_df_melted, aes(x = `1/[S] (1/mM)`, y = value)) +
    geom_hline(yintercept = 0.0, colour = "black") +
    geom_vline(xintercept = 0.0, colour = "black") + geom_point() +
    xlim((-1 / answers[3]),LB_df[1,1]) +
    geom_abline(intercept = 1 / answers[1], 
                slope = answers[3] / answers[1], colour = "red") +
    geom_abline(intercept = 1 / answers[2], 
                slope = answers[4] / answers[2], colour = "blue") +
    labs(x = "1/[S] (1/mM)", y = "1/v (min/mmol)") +
    ggtitle("Lineweaver-Burk Plot") + theme(plot.title = element_text(size = 28,face = "bold")) +
    stat_poly_eq(data = LB_df, formula = y ~ x, 
                 aes(y = `1/v (min/mmol)`, label = paste(..eq.label.., ..rr.label.., sep = "~~~")), 
                 parse = TRUE, color = "red", label.y = 0.95) +
    stat_poly_eq(data = LB_df, formula = y ~ x, 
                 aes(y = `1/v_inh (min/mmol)`, label = paste(..eq.label.., ..rr.label.., sep = "~~~")), 
                 parse = TRUE, color = "blue", label.y = 0.9)
  
  
  # create EH data
  vel_div_conc_S <- vel / conc_S
  vel_inh_div_conc_S <- vel_inh / conc_S
  EH_data <- round(cbind(vel_div_conc_S, vel, vel_inh_div_conc_S, vel_inh), 3)
  colnames(EH_data) <- c("v/[S] (L/min)", "v (mmol/min)", "v_inh/[S] (L/min)", "v_inh (mmol/min)")
  EH_df <- as.data.frame(EH_data, optional = TRUE)
  
  # plot EH plot
  plot_EH <- ggplot() + geom_point(data = EH_df[1:2], aes(x = `v/[S] (L/min)`, y = `v (mmol/min)`)) +
    geom_hline(yintercept = 0.0, colour = "black") +
    geom_vline(xintercept = 0.0, colour = "black") + geom_point() +
    xlim(0, answers[1] /answers[3]) + ylim(0, answers[1]) +
    geom_abline(intercept = answers[1], 
                slope = - answers[3], colour = "red") +
    geom_point(data = EH_df[3:4], aes(x = `v_inh/[S] (L/min)`, y = `v_inh (mmol/min)`)) +
    geom_abline(intercept = answers[2], slope = - answers[4], colour = "blue") +
    labs(x = "v/[S] (L/min)", y = "v (mmol/min)") +
    ggtitle("Eadie-Hofstee Plot") + theme(plot.title = element_text(size = 28,face = "bold")) +
    stat_poly_eq(data = EH_df, formula = y ~ x, 
                 aes(x = `v/[S] (L/min)`, y = `v (mmol/min)`, label = paste(..eq.label.., ..rr.label.., sep = "~~~")), 
                 parse = TRUE, color = "red", label.y = 0.95, label.x = 0.95) +
    stat_poly_eq(data = EH_df, formula = y ~ x, 
                 aes(x = `v_inh/[S] (L/min)`, y = `v_inh (mmol/min)`, label = paste(..eq.label.., ..rr.label.., sep = "~~~")), 
                 parse = TRUE, color = "blue", label.y = 0.9, label.x = 0.95)
  
    
  # create HW data
  conc_S_div_vel <- conc_S / vel
  conc_S_div_vel_inh <- conc_S / vel_inh
  HW_data <- round(cbind(conc_S, conc_S_div_vel, conc_S_div_vel_inh), 3)
  colnames(HW_data) <- c("[S] (mM)", "[S]/v (min/L)", "[S]/v_inh (min/L)")
  HW_df <- as.data.frame(HW_data, optional = TRUE)
  HW_df_melted <- reshape2::melt(HW_df, id = "[S] (mM)")
  
  # create HW plot
  plot_HW <- ggplot(data = HW_df_melted, aes(x = `[S] (mM)`, y = value)) +
    geom_hline(yintercept = 0.0, colour = "black") +
    geom_vline(xintercept = 0.0, colour = "black") + geom_point() +
    xlim(- answers[4], max(HW_df[,1])) +
    geom_abline(intercept = answers[3] / answers[1], 
                slope = 1 / answers[1], colour = "red") +
    geom_abline(intercept = answers[4] / answers[2], 
                slope = 1 / answers[2], colour = "blue") +
    labs(x = "[S] (mM)", y = "[S]/v (min/L)") +
    ggtitle("Hanes-Woolf Plot") + theme(plot.title = element_text(size = 28,face = "bold")) +
    stat_poly_eq(data = HW_df, formula = y ~ x, 
                 aes(y = `[S]/v (min/L)`, label = paste(..eq.label.., ..rr.label.., sep = "~~~")), 
                 parse = TRUE, color = "red", label.y = 0.95) +
    stat_poly_eq(data = HW_df, formula = y ~ x, 
                 aes(y = `[S]/v_inh (min/L)`, label = paste(..eq.label.., ..rr.label.., sep = "~~~")), 
                 parse = TRUE, color = "blue", label.y = 0.9)
  
  
  # combine MM_data and answers in list
  list_dfs <- list(answers_df, MM_data_df, LB_df, plot_LB, EH_df, plot_EH, HW_df, plot_HW, plot_MM)
  return(list_dfs)
}

create_data()

# multiply function to create different data sets 
test_df <- lapply(1:4, create_data)
#test_df


create_files <- function(test_df) {
  #location of files
  save_path <- paste0(work_path, "/MM_data/Noncomp")
  dir.create(file.path(save_path, "Data"))
  dir.create((file.path(save_path, "Answers")))
  save_data <- file.path(save_path, "Data/")
  save_answers <- file.path(save_path, "Answers/")
  
  #create files with data
  data_list <- list()
  for(i in test_df) {
    data_list <- append(data_list, i[2])
  }

  # write files with data + extras ('invulformulier')  
  names(data_list) <- paste0("dataset", 51:54) #adjust for number of comp.data sets
  for(s in seq_along(names(data_list))) {
    data_file_name = paste0(save_data, names(data_list[s]), ".xlsx")
    wb <- openxlsx::loadWorkbook(paste0(save_path, "/template_MM_noncomp.xlsx"))
    
    name_sheet <- paste0("Dataset", 50 + s) #adjust for number of comp.data sets
    renameWorksheet(wb, 1, name_sheet)
    writeData(wb, 1, name_sheet)
    writeData(wb, 1, data_list[[s]], startRow = 2)
    protectWorksheet(wb, 1)
    writeData(wb, 2, name_sheet)
    writeData(wb, 2, data_list[[s]], startRow = 2)
    writeData(wb, 3, name_sheet)
    writeData(wb, 3, data_list[[s]], startRow = 2)
    openxlsx::saveWorkbook(wb, data_file_name, overwrite = TRUE)
  }
  
  # make a list with answers to write to files
  answers_list <- list()
  original_data_list <- list()
  graph_MM_list <- list()
  datapoints_LB <- list()
  graph_LB_list <- list()
  datapoints_EH <- list()
  graph_EH_list <- list()
  datapoints_HW <- list()
  graph_HW_list <- list()

  for(i in test_df) {
    answers_list <- append(answers_list, i[1]) #answers
    all_answers <- rbindlist(answers_list)
    original_data_list <- append(original_data_list, i[2]) #original data
    graph_MM_list <- append(graph_MM_list, i[9]) #original (MM) graph
    datapoints_LB <- append(datapoints_LB, i[3]) #LB data
    graph_LB_list <- append(graph_LB_list, i[4]) #LB graph
    datapoints_EH <- append(datapoints_EH, i[5]) #EH data
    graph_EH_list <- append(graph_EH_list, i[6]) #EH graph
    datapoints_HW <- append(datapoints_HW, i[7]) #HW data
    graph_HW_list <- append(graph_HW_list, i[8]) #HW graph
  }

  # write file with answers to all datasets
  write.xlsx2(all_answers, paste0(save_answers, "all_answers_noncomp.xlsx"),
              col.names = TRUE)
  
  
  # make plots from the data sets
  for(c in 1:length(graph_MM_list)) {
    file_MM_png = paste0(save_path, "MM_plot", c, ".png")
    png(file_MM_png)
    print(graph_MM_list[[c]])
    dev.off()
  }
  
  for(n in 1:length(graph_LB_list)) {
    file_LB_png = paste0(save_path, "LB_plot", n, ".png")
    png(file_LB_png)
    print(graph_LB_list[[n]])
    dev.off()
  }
  
  for(a in 1:length(graph_EH_list)) {
    file_EH_png = paste0(save_path, "EH_plot", a, ".png")
    png(file_EH_png)
    print(graph_EH_list[[a]])
    dev.off()
  }
  
  for(b in 1:length(graph_HW_list)) {
    file_HW_png = paste0(save_path, "HW_plot", b, ".png")
    png(file_HW_png)
    print(graph_HW_list[[b]])
    dev.off()
  }
 
  
  names(answers_list) <- paste0("answers", 51:54) #adjust for number of comp.data sets
  
  # writing data, answers and plots in 1 file (for all data sets)
  for(k in seq_along(names(answers_list))) {
    file_name = paste0(save_answers, names(answers_list[k]), ".xlsx")
    file_MM_png = paste0(save_path, "MM_plot", k, ".png")
    file_LB_png = paste0(save_path, "LB_plot", k, ".png")
    file_EH_png = paste0(save_path, "EH_plot", k, ".png")
    file_HW_png = paste0(save_path, "HW_plot", k, ".png")
    wb <- openxlsx::createWorkbook(file_name)
    
    addWorksheet(wb, "Michaelis-Menten")
    writeData(wb, "Michaelis-Menten", original_data_list[[k]])
    writeData(wb, "Michaelis-Menten", answers_list[[k]], startRow = 10)
    insertImage(wb, "Michaelis-Menten", file_MM_png, width = 6, height = 6, 
                startCol = 6)
    
    addWorksheet(wb, "Lineweaver-Burk")
    writeData(wb, "Lineweaver-Burk", original_data_list[[k]])
    writeData(wb, "Lineweaver-Burk", datapoints_LB[[k]], startRow = 10)
    writeData(wb, "Lineweaver-Burk", answers_list[[k]], startRow = 19)
    insertImage(wb, "Lineweaver-Burk", file_LB_png, width = 6, height = 6, 
                startCol = 6)
    
    addWorksheet(wb, "Eadie-Hofstee")
    writeData(wb, "Eadie-Hofstee", original_data_list[[k]])
    writeData(wb, "Eadie-Hofstee", datapoints_EH[[k]], startRow = 10)
    writeData(wb, "Eadie-Hofstee", answers_list[[k]], startRow = 19)
    insertImage(wb, "Eadie-Hofstee", file_EH_png, width = 6, height = 6, 
                startCol = 6)
    
    addWorksheet(wb, "Hanes-Woolf")
    writeData(wb, "Hanes-Woolf", original_data_list[[k]])
    writeData(wb, "Hanes-Woolf", datapoints_HW[[k]], startRow = 10)
    writeData(wb, "Hanes-Woolf", answers_list[[k]], startRow = 19)
    insertImage(wb, "Hanes-Woolf", file_HW_png, width = 6, height = 6, 
                startCol = 6)
    
    openxlsx::saveWorkbook(wb, file_name, overwrite = TRUE)
    unlink(file_MM_png)
    unlink(file_LB_png)
    unlink(file_EH_png)
    unlink(file_HW_png)
  }  
  
  
}


create_files(test_df)
   