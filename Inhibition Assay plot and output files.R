library("readxl")
library("writexl")
library(rio)
library(xlsx)
library(dplyr)
library(ggplot2)
library(RColorBrewer)



time=c(0, 3, 23,27,44, 50 )  #classic smeg
#time=c(0,19,26,45,50)    #trigger smeg
#time=c(0,44,75,139)     #trigger marinum
#time=c(0, 49, 117, 165, 213) #classic marinum


strain=c("7-1861-15","MSH 2a1","4-1861-60","3-1861-45","10-1861-75","1-1861-105","9-1861-50","3-2157-15","8-1861-60","MSH 2a1*","3.1-1861-15","2-1861-30","3-1861-15","2-2157-15","5-1861-230","7-1861-30","8-1861-30", "S.bikiniensis", "ctrl 1", "ctrl 2")


#Title <- expression(paste("Influence of ", italic("Actinobacteria"), " extracts (aeration, 9 day culture) on growth of ", italic("M. smegmatis")))
#Subtitle= "cell/biomass extracts"
#Experiment= paste("9 days Cells", "Smegmatis")

k=1   #номер чашки

for (k in 1:4) {

  h=16+k   #номер строки в exсel
  
  
  leave_ctrl_name= "ctrl 2"
  #rename_ctrl_name= "Methanol"  #переименовать контроль
  
  
  remove_ctrl_name= "ctrl 1"  #удалить лишний контроль
  remove_ctrl= TRUE
  
  
  folder_path= "D:/input_data" #папка входных данных
  sheet= "Sheet 1"
  n0=2; #эксперимент №1
  n1=3  #эксперимент №2
  
  
  
  
  datapath= file.path(folder_path,paste("Expriment Conditions Names",".xlsx", sep=""))
  dataset_condition_names= read_excel(datapath, sheet = "Sheet1")
  
  
  Experiment= dataset_condition_names[h,2]
  a= as.character(dataset_condition_names[h,3])
  b= as.character(dataset_condition_names[h,4])
  c= as.character(dataset_condition_names[h,5])
  d= as.character(dataset_condition_names[h,6])
  
  
  
  Title <- bquote(.(a)~italic(.(b))~.(c)~italic(.(d)))
  
  
  Subtitle = dataset_condition_names[h,7]
  rename_ctrl_name= dataset_condition_names[h,8]
  
  
  for (i in k:k) {
    bunch_1= paste(n0,"_plate_", i, sep="")
    datapath= file.path(folder_path,paste(bunch_1,".xlsx", sep=""))
    dataset_1= read_excel(datapath, sheet = sheet)
    
    
    bunch_2= paste(n1,"_plate_", i, sep="")
    datapath= file.path(folder_path,paste(bunch_2,".xlsx", sep=""))
    dataset_2= read_excel(datapath, sheet = sheet)
    
  }
  
  
  col_range= list(2:4, 5:7, 8:10, 11)
  row_l1= list(7,8,9,10,11,12, 7:9, 10:12)
  row_l2= list(21,22,23,24,25,26, 21:23, 24:26)
  row_l3= list(35,36,37,38,39,40, 35:37, 38:40)
  row_l4= list(49,50,51,52,53,54, 49:51, 52:54)
  row_l5= list(63,64,65,66,67,68, 63:65, 66:68)
  row_rang= list(row_l1, row_l2, row_l3, row_l4, row_l5)
  
  
  binder <- function(dataset, iterator) {
    s=1
    c=1
    r=1
    q1=NULL
    for (i in 1:18) {
      if(r==7) {
        c=c+1
        r=1
      }
      q=dataset[as.numeric(gsub(":.*","",row_rang[[iterator]][r])):as.numeric(gsub(".*:","",row_rang[[iterator]][r])),
                  as.numeric(gsub(":.*","",col_range[c])):as.numeric(gsub(".*:","",col_range[c]))]
      q=cbind(s, q)
      colnames(q)[1]="Strain"
      colnames(q)[2]=1
      colnames(q)[3]=2
      colnames(q)[4]=3
      q1= rbind(q1,q)
      s=s+1
      r=r+1
    }
    
    r=7
    c=4
    s0=1
    for (i in 1:2) {
      s="ctrl"
      
      q=dataset[as.numeric(gsub(":.*","",row_rang[[iterator]][r])):as.numeric(gsub(".*:","",row_rang[[iterator]][r])),
                  as.numeric(gsub(":.*","",col_range[c])):as.numeric(gsub(".*:","",col_range[c]))]
      q=t(q)
      
      s=paste(s, s0)
      q=cbind(s, q)
      colnames(q)[1]="Strain"
      colnames(q)[2]=1
      colnames(q)[3]=2
      colnames(q)[4]=3
      q1= rbind(q1,q)
      s0=s0+1
      r=r+1
    }
    
    return(q1)
  }
  
  
  
  i=2
  final_output=NULL
  first_output= NULL
  output=NULL
  output_subtract=NULL
  for (i in 1:5) {
    print(i)
    first_measure=binder(dataset_1, i)
    second_measure=binder(dataset_2, i)
    second_measure= second_measure[-c(1)]
    if (i==1) {
      first_output=cbind(first_measure, second_measure)
      first_output= data.frame(apply(first_output, 2, function(x) as.numeric(as.character(x))))
    } else {
      output= cbind(first_measure, second_measure)
      output= data.frame(apply(output, 2, function(x) as.numeric(as.character(x))))
      output=output-first_output
      output=rbind("", output)
      
    }
    final_output= rbind(final_output, output)
  }
  final_output$Strain= NULL
  final_output= data.frame(apply(final_output, 2, function(x) as.numeric(as.character(x))))
  
  zero_data= data.frame(matrix(0, 20, 6)) #создаю таблицу с нулями
  colnames(zero_data) <- colnames(final_output)
  
  final_output= rbind(zero_data,final_output)
  
  
  mean_data=data.frame(rowMeans(final_output)) #получаю среднее
  colnames(mean_data) = "mean"
  stdev_data= data.frame(apply(final_output, 1, sd)) #получаю отклоненние
  colnames(stdev_data) = "sd"
  pre_plot_data= cbind(mean_data, stdev_data)
  pre_plot_data <- pre_plot_data[!apply(is.na(pre_plot_data) | pre_plot_data == "", 1, all),] #удаляю все пустые строуки и NA
  pre_plot_data= cbind("", pre_plot_data)
  m=1
  for (n in 1:5) {
    for (i in 1:20) {
      
        pre_plot_data[m,1]= strain[i]
      
      m=m+1
    }
  }
  
  
  #pre_plot_data= data.frame(apply(pre_plot_data, 2, function(x) as.numeric(as.character(x))))
  
  colnames(pre_plot_data)[1]= "Strain"
  pre_plot_data= cbind(pre_plot_data, "")
  colnames(pre_plot_data)[4]= "Time"
  pre_plot_data[,4]
  m=1
  for (n in 1:5) {
    for (i in 1:20) {
      if (n==1) {
        pre_plot_data[m,4]= time[1] 
      } else if (n==2) {
        pre_plot_data[m,4]= time[2]
      } else if (n==3) {
        pre_plot_data[m,4]= time[3]
      }else if (n==4){
        pre_plot_data[m,4]= time[4]
      } else if (n==5){
        pre_plot_data[m,4]= time[5]
      }
      m=m+1
    }
  }
  pre_plot_data[,4] <- as.numeric(as.character(pre_plot_data[,4]))
  strain_plot_data= pre_plot_data[- grep("ctrl", pre_plot_data$Strain),]
  #strain_plot_data= strain_plot_data[- grep("bikiniensis", strain_plot_data$Strain),]
  
  ctrl_plot_data <- pre_plot_data[grep("ctrl", pre_plot_data$Strain),]
  
  
  if (remove_ctrl==TRUE) {
  ctrl_plot_data <- ctrl_plot_data[!grepl(remove_ctrl_name, ctrl_plot_data$Strain),]
  }
  
  ctrl_plot_data$Strain <- gsub(leave_ctrl_name, rename_ctrl_name, ctrl_plot_data$Strain)
  
  
  #pre_plot_data$Strain
  nb.cols <- 20
  mycolors <- colorRampPalette(brewer.pal(8, "Set2"))(nb.cols)
  
  p=ggplot(strain_plot_data, aes(x=Time, y=mean, group= Strain, color= Strain)) + geom_line(size=0.9) + geom_point(size=1.6)+
    geom_line(data=ctrl_plot_data, linetype = "dashed", size=0.9)+ geom_point(data=ctrl_plot_data, size=1.6)+ geom_errorbar(data=ctrl_plot_data, aes(ymin=mean-sd, ymax=mean+sd), width=0.4) +
    labs(subtitle=paste("Title"), 
         y="OD", 
         x="Time, hours",
         
    )+
    #gghighlight(Strain == list("ctrl 1", "S.bikiniensis"))
                #unhighlighted_params = list(colour = "grey70", alpha = 0.7))+
    scale_fill_manual(values = mycolors)+
    geom_errorbar(aes(ymin=mean-sd, ymax=mean+sd), width=0.4) +
    
    theme(legend.text=element_text(size=7))+
    theme_minimal()
    #geom_point(data = subset(pre_plot_data,  Strain == 1 & Other ==2),color = colors[3])+
    #geom_line(data = subset(pre_plot_data,  Strain =="ctrl 1"),color = "grey10", linetype = "dashed")
    #ylim(0,0.3)+
  p
  
  
  
  
  p1=ggplot(strain_plot_data, aes(x=Time, y=mean, group= Strain, color= Strain)) + geom_line(size=0.3) + geom_point(size=0.5)+
    geom_line(data=ctrl_plot_data, linetype = "dashed", size=0.3)+ geom_point(data=ctrl_plot_data, size=0.5)+ geom_errorbar(data=ctrl_plot_data, aes(ymin=mean-sd, ymax=mean+sd), width=0.4) +
    labs(title=Title,
        subtitle=paste(Subtitle), 
         y=paste("OD"), 
         x="Time, hours",
         
    )+
    #gghighlight(Strain == list("ctrl 1", "S.bikiniensis"))
    #unhighlighted_params = list(colour = "grey70", alpha = 0.7))+
    scale_fill_manual(values = mycolors)+
    geom_errorbar(aes(ymin=mean-sd, ymax=mean+sd), width=0.3) +
    theme_minimal()+
    theme(legend.text=element_text(size=5, color= "grey30"),
          legend.title = element_text(size=6, color= "grey30"),
          plot.title = element_text(size=5, color= "grey30"),
          plot.subtitle = element_text(size=5, color= "grey30"),
          axis.text.x = element_text(size = 5),
          axis.text.y = element_text(size = 5),  
          axis.title.x = element_text(size = 6),
          axis.title.y=element_text(size = 6),
          axis.title=element_text(size=9),
          legend.key.size = unit(0.22, "cm")
          )
    
  #geom_point(data = subset(pre_plot_data,  Strain == 1 & Other ==2),color = colors[3])+
  #geom_line(data = subset(pre_plot_data,  Strain =="ctrl 1"),color = "grey10", linetype = "dashed")
  #ylim(0,0.3)+
  p1 
  
  #подготовка файлп на t-тест
  k=84
  p_value_dataset=NULL
  for (i in 1:20) {
    rm(t_test_dataset)
    t_test_dataset=final_output[k+i,]
    t_test_dataset= rbind(list(1,1,1,2,2,2), t_test_dataset)
    t_test_dataset=t(t_test_dataset)
    colnames(t_test_dataset)=c("experiment", as.character(strain[i]))
    p_value <- t.test(t_test_dataset[,1], t_test_dataset[,2], paired=TRUE)$p.value
    p_value=data.frame(p_value)
    p_value=cbind(strain[i], p_value)
    colnames(p_value)[1]= "Strain"
    p_value_dataset= rbind(p_value_dataset, p_value)
  }
  
  #подготовка файла на PCA модель
  pca_dataset=NULL
  #pca_dataset=final_output[85:104,]
  pca_dataset=final_output[64:83,]
  pca_dataset=cbind(strain, pca_dataset)
  colnames(pca_dataset)= c("Strain","1.1","1.2","1.3", "2.1","2.2","2.3" )
  pca_dataset=cbind("", pca_dataset)
  pca_dataset[1]=Subtitle
  colnames(pca_dataset)[1]= "Fraction"
  pca_dataset=cbind("", pca_dataset)
  pca_dataset[1]=d
  colnames(pca_dataset)[1]= "Test strain"
  pca_dataset=cbind("", pca_dataset)
  pca_dataset[1]=dataset_condition_names[h,1]
  colnames(pca_dataset)[1]= "Time"
  pca_dataset=cbind("", pca_dataset)
  pca_dataset[1]=b
  colnames(pca_dataset)[1]= "Type"
  pca_dataset=cbind("", pca_dataset)
  pca_dataset[1]=Experiment
  colnames(pca_dataset)[1]= "Primary Sample"
  
  
  filepath=paste(folder_path,"/",Experiment,"_p-value",".xlsx", sep="") #экспорт p-value файла
  export(p_value_dataset, filepath)
  
  filepath=paste(folder_path,"/",Experiment,"_PCA_data",".xlsx", sep="") #экспорт PCF файла
  export(pca_dataset, filepath)
  
  filepath=paste(folder_path,"/",Experiment,"_data_plot", ".xlsx", sep="") #экспорт общего файла
  export(pre_plot_data, filepath)
  
  ggsave(p1, file=paste(folder_path,"/",Experiment,".jpg", sep=""), width = 11, height = 6, units = "cm") #экспорт картинки
  print("Done")

}

