# Hello, world!
#
# This is an example function named 'hello'
# which prints 'Hello, world!'.
#
# You can learn more about package authoring with RStudio at:
#
#   http://r-pkgs.had.co.nz/
#
# Some useful keyboard shortcuts for package authoring:
#
#   Install Package:           'Ctrl + Shift + B'
#   Check Package:             'Ctrl + Shift + E'
#   Test Package:              'Ctrl + Shift + T'

keyword_function <- function(xlsxfile, keyword, outname){
  library(readxl)
  sheetnames <- excel_sheets(xlsxfile)
  Alldata <- c()
  for(i in 1:length(sheetnames)){
    data0 <- read_excel(xlsxfile,sheet = i)[-1:-3,]
    Alldata <- rbind(Alldata,data0)
  }

  data <- unique(as.matrix(Alldata[which(Alldata[,1]!="主管簽核：1."),]))

  for(i in 1:length(keyword)){
    data <- data[grep(keyword[i], data[,10]),]
  }

  filename <- paste(outname,".csv")
  data <- data.frame("拜訪日期" = data[,1], "星期" = data[,2], "業種" = data[,3],
                     "級別" = data[,4], "地區" = data[,5], "配合人員" = data[,6],
                     "客戶號碼" = data[,7], "客戶名稱" = data[,8],
                     "接洽人" = data[,9], "訪問內容" = data[,10],
                     "競爭同業" = data[,11], "競爭商品" = data[,12])
  write.csv(data,filename)
}
