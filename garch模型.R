rm(list=ls())
## 读取Excel数据
# XLConnect包，需要JAVA环境
install.packages("XLConnect") # 需要rJava包
library(XLConnect)
# 读取文件（两步法）
x <- loadWorkbook("/Users/nan-fan/Desktop/0-数字货币波动性研究/数字货币数据处理/acd3.xlsx") # 导入文档
coindata_all1 <- readWorksheet(x, 1) # 读取文档 # 不同的工作表，数字为工作表编号
head(coindata_all1) # 查看数据，显示前6行

is.data.frame(coindata_all1) # 用来判断数据类型



## 进行平稳性检验
# install.packages("rmgarch")
# library(rmgarch)
# install.packages("tseries")
library(tseries)
adf.test(coindata_all1$btc_p) # ADF检验，使用命令进行单位根检验，判断：p 值远小于 0.01 说明拒绝原假设，即序列是平稳的。

adf.test(coindata_all1[,3])
for (i in 2:291) {adf.test(coindata_all1[,i])}
warnings()

install.packages("forecast")
library(forecast)
md1 <- auto.arima(coindata_all1[,2])
md1

# 系数显著性检验
# install.packages("knitr") # 启用kable命令
library(knitr)
t <- md1$coef / sqrt(diag(md1$var.coef))
p_1 <- 2 * (1 - pnorm(abs(t)))
kable(rbind(p = p_1), caption = "Result of Coef. Test")

for (i in 2:2) {auto.arima(coindata_all1[,i])}
show(auto.arima())
for (i in 2:291) {auto.arima(coindata_all1[,i])}

# 拟合疏系数arma
mean_md_1 <- Arima(r.data,
                   order = c(4, 0, 4),
                   fixed = c(NA, 0, 0, NA, NA, 0, 0, NA, 0))

# 疏系数模型的系数显著性检验
t <- mean_md_1$coef / sqrt(diag(mean_md_1$var.coef))
p_2 <- 2 * (1 - pnorm(abs(t)))
kable(rbind(p = p_2[p_2 != 1]), caption = "Result of Coef. Test")

# ARCH效应检验（原假设是没有ARCH效应）
print(paste0('m = 10'))
archTest(mean_md_1$res, lag = 10)
print(paste0('m = 20'))
archTest(mean_md_1$res, lag = 20)

# 使用tseries拟合sGARCH（只能拟合方差模型，没有考虑均值模型）
out <- garch(coindata_all1[,2], order = c(1, 1))
summary(out)

# 模型优化
# 模型选择
md_name <- c("sGARCH", "eGARCH", "gjrGARCH",
             "apARCH", "TGARCH", "NAGARCH")
md_dist <- c("norm", "std", "sstd", "ged", "jsu")

# 设定函数
model_train <- function(data_, m_name, m_dist) {
  res <- NULL
  for (i in m_name) {
    sub_model <- NULL
    main_model <- i
    if (i == "TGARCH" | i == "NAGARCH") {
      sub_model <- i
      main_model <- "fGARCH"
    }
    
    for (j in m_dist) {
      model_name <- paste0('ARMA(4,4)', '-',
                           i, '-', j,
                           sep = "")
      if (main_model == "fGARCH") {
        model_name <- paste0('ARMA(4,4)', '-',
                             sub_model, '-', j,
                             sep = "")
      }
      print(model_name)
      # 模型设定
      mean.spec <- list(armaOrder = c(4, 4), include.mean = F,
                        archm = F, archpow = 1, arfima = F,
                        external.regressors = NULL)
      var.spec <- list(model = main_model, garchOrder = c(1, 1),
                       submodel = sub_model,
                       external.regressors = NULL,
                       variance.targeting = F)
      dist.spec <- j
      my_spec <- ugarchspec(mean.model = mean.spec,
                            variance.model = var.spec,
                            distribution.model = dist.spec)
      # 模型拟合
      my_fit <- ugarchfit(data = data_, spec = my_spec)
      res <- rbind(res, c(list(ModelName = model_name),
                          list(LogL = likelihood(my_fit)),
                          list(AIC = infocriteria(my_fit)[1]),
                          list(BIC = infocriteria(my_fit)[2]),
                          list(RMSE = sqrt(mean(residuals(my_fit)^2))),
                          list(md = my_fit)))
    }
  }
  return(res)
}

# 模型拟合
my_model <- model_train(coindata_all1[,2], md_name, md_dist)

# 保存模型结果
write.table(my_model[, 1:5],
            "/Users/nan-fan/Desktop/garch_test/result.txt",
            sep = " ")

# 读取模型结果
res_table <- read.table("/Users/nan-fan/Desktop/garch_test/result.txt",
                        header = TRUE,
                        sep = " ")
kable(res_table, caption = "Result of GARCH Model", align = 'l')

# 选择模型再拟合
md_name <- c("eGARCH")
md_dist <- c("norm", "snorm", "std", "sstd", "ged", "sged", "jsu")
new_model <- model_train(coindata_all1[,2], md_name, md_dist)

# 信息准则表
my_info_cri <- NULL
for (i in 1:dim(new_model)[1]) {
  new_aic <- round(infocriteria(new_model[i, 6]$md)[1], digits = 4)
  new_bic <- round(infocriteria(new_model[i, 6]$md)[2], digits = 4)
  new_sib <- round(infocriteria(new_model[i, 6]$md)[3], digits = 4)
  new_hq <- round(infocriteria(new_model[i, 6]$md)[4], digits = 4)
  new_llk <- round(likelihood(new_model[i, 6]$md), digits = 4)
  my_info_cri <- rbind(my_info_cri, c(new_model[i, 1],
                                      Akaike = new_aic,
                                      Bayes = new_bic,
                                      Shibata = new_sib,
                                      HQ = new_hq,
                                      LLH = new_llk))
}


# 模型系数显著性表
my_cof_table <- NULL
my_cof_names <- NULL
for (i in 1:dim(new_model)[1]) {
  my_cof_names <- rbind(my_cof_names, new_model[i, 1]$ModelName)
  my_df <- data.frame(t(round(new_model[i, 6]$md@fit$robust.matcoef[, 4],
                              digits = 4)))
  if (i == 1) {
    my_cof_table <- data.frame(my_df)
  }
  else {
    my_cof_table <- full_join(my_cof_table, my_df)
  }
}
row.names(my_cof_table) <- my_cof_names[, 1]
my_cof_table <- t(my_cof_table)


# 参数稳定性个别检验表
my_nyblom_table <- NULL
my_nyblom_name <- NULL
for (i in 1:dim(new_model)[1]) {
  my_nyblom_name <- rbind(my_nyblom_name,
                          new_model[i, 1]$ModelName)
  my_df <- data.frame(t(round(nyblom(new_model[i, 6]$md)$IndividualStat,
                              digits = 4)))
  if (i == 1) {
    my_nyblom_table <- data.frame(my_df)
  }
  else {
    my_nyblom_table <- full_join(my_nyblom_table, my_df)
  }
}
row.names(my_nyblom_table) <- my_nyblom_name[, 1]
my_IC <- NULL
for (i in 1:dim(new_model)[1]) {
  my_IC <- cbind(my_IC, nyblom(new_model[1, 6]$md)$IndividualCritical)
}
my_nyblom_table <- rbind(t(my_nyblom_table), my_IC)

# 参数稳定性联合检验表
my_nyblom_table2 <- NULL
my_nyblom_name2 <- NULL
for (i in 1:dim(new_model)[1]) {
  my_nyblom_name2 <- rbind(my_nyblom_name2,
                           new_model[i, 1]$ModelName)
  df1 <- data.frame(JoinStat = round(nyblom(new_model[i,6]$md)$JointStat,
                                     digits = 4))
  df2 <- data.frame(mm = round(nyblom(new_model[i, 6]$md)$JointCritical,
                               digits = 4))
  temp_table <- cbind(df1, t(df2))
  if (i == 1) {
    my_nyblom_table2 <- temp_table
  }
  else {
    my_nyblom_table2 <- full_join(my_nyblom_table2, temp_table)
  }
}
row.names(my_nyblom_table2) <- my_nyblom_name2

# 模型分布拟合度p值检验表
my_gof_table <- NULL
my_gof_name <- NULL
group_name <- c(20, 30, 40, 50)
for (i in 1:dim(new_model)[1]) {
  my_gof_name <- rbind(my_gof_name,
                       new_model[i, 1]$ModelName)
  df1 <- round(gof(new_model[i, 6]$md,
                   groups = group_name)[1:4, 3],
               digits = 4)
  if (i == 1) {
    my_gof_table <- df1
  }
  else {
    my_gof_table <- rbind(my_gof_table, df1)
  }
}
row.names(my_gof_table) <- my_gof_name
colnames(my_gof_table) <- c(20, 30, 40, 50)

# 符号偏误显著性检验表
my_sign_table <- NULL
my_sign_name <- NULL
row_sign_name <- row.names(signbias(new_model[1, 6]$md))
for (i in 1:dim(new_model)[1]) {
  my_sign_name <- rbind(my_sign_name,
                        new_model[i, 1]$ModelName)
  df1 <- round(signbias(new_model[i, 6]$md)$prob, digits = 4)
  if (i == 1) {
    my_sign_table <- df1
  }
  else {
    my_sign_table <- rbind(my_sign_table, df1)
  }
}
row.names(my_sign_table) <- my_sign_name
colnames(my_sign_table) <- row_sign_name

# 展示表格
kable(my_cof_table,
      caption = "P-value Table of Coefficients")
kable(my_nyblom_table,
      caption = "P-value Table of Nyblom Individual Stability Test")
kable(my_nyblom_table2,
      caption = "P-value Table of Nyblom Joint Stability Test")
kable(my_sign_table,
      caption = "P-value Table of Sign Bias Test")
kable(my_gof_table,
      caption = "P-value Table of Goodness-of-fit")
kable(my_info_cri,
      caption = "Table of Information Criteria")

# 最终模型系数表
kable(new_model[6, 6]$md@fit$robust.matcoef,
      caption = "Table of Final Model Coef.")


## 建立方差模型
# ARCH模型
# 对残差进行ARCH效应检验
# 原假设是序列不存在自相关，在残差的平方序列中可以检验条件异方差。
install.packages("MTS")
library(MTS)
archTest(coindata_all1[,2])
archTest(coindata_all1[,3])

# 建立GARCH模型
library(tseries)
garch(coindata_all1[,2])






