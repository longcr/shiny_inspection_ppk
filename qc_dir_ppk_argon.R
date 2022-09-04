# read QC DIR and output process performance
# 2020-05-09
# Cliff Long


# DATA FLOW ###################################################################

# web page to upload data from DIR
# upload

# web page


# LOAD PACKAGES ###############################################################

library(readxl)
library(janitor)
library(dplyr)
library(tidyr)
library(ggplot2)
library(qcc)
library(ggQC)
library(purrr)
library(broom)
library(SixSigma)
#library(qualityTools)
library(qqplotr)


# LOAD FUNCTIONS ##############################################################

# function 1 ------------------------------------------------------------------
fn_capability_plot <- function(fdatcp, fspecscp){
  
  require(ggplot2)
  
  
  # TEST
  # fdatcp <- d_dat_nestj$data[[2]]
  # fspecscp <- unlist(d_dat_nestj$speclist[[2]])

  fspecscp <- unlist(fspecscp)
  
  fnspecscp_nom <- fspecscp[1]
  fnspecscp_lsl <- fspecscp[2] 
  fnspecscp_usl <- fspecscp[3]

  
  # forces response variable name to "value"
  fdatcp <- fdatcp %>% rename_at(2, ~"value")
  
  
  # number of bins
  # binw <- max(6, ceiling(sqrt(length(fdatcp$value))))
  
  
  # create initial histogram of data
  cp_plot <- fdatcp %>% 
    ggplot(aes(x = value)) + 
    geom_histogram(aes(y = ..density..), 
                   # bins = binw, 
                   fill = 'gray', 
                   color = 'black')
  
  
  # create density curve data
  xmean <- mean(fdatcp$value)
  xsd <- sd(fdatcp$value)
  xmin <- xmean - 4*xsd
  xmax <- xmean + 4*xsd
  xden <- seq(xmin, xmax, length.out=100)
  fdatcp_dist <- data.frame(x = xden, y = dnorm(xden, xmean, xsd))
  
  
  # test density
  # ggplot() + geom_line(data = fdatcp_dist, aes(x = xden, y = y), color = "red") 

  
  # add dist to histogram
  cp_plot <- cp_plot + 
    geom_line(data = fdatcp_dist, aes(x = xden, y = y), color = "red") 

    
  # add spec limits to histogram
  cp_plot <- cp_plot + 
    geom_vline(xintercept = c(fnspecscp_lsl, fnspecscp_usl), linetype = 2, color = 'red')

  return(cp_plot)
  
}



# function 2 ------------------------------------------------------------------
fn_ppk <- function(fndatppk, fnspecppk){
  
  # TEST
  # fndatppk <- d_dat_nestj$data[[1]]
  # fnspecppk <- unlist(d_dat_nestj$speclist[[2]])
  
  
  fnspecppk <- unlist(fnspecppk)
  
  fnspecppk_nom <- fnspecppk[1]
  fnspecppk_lsl <- fnspecppk[2] 
  fnspecppk_usl <- fnspecppk[3]
  
  
  # the cp function won't accept dimensioning where nominal is outside of LSL, USL
  if ((fnspecppk_nom < fnspecppk_lsl) | (fnspecppk_nom > fnspecppk_usl)){
    fnspecppk_nom = mean(c(fnspecppk_lsl, fnspecppk_usl))
  }
  

  # forces response variable name to "value"
  fndatppk <- fndatppk %>% rename_at(2, ~"value")
  

  # ppkout <- cp(x = fndatppk$value, 
  #              distribution = "normal",
  #              lsl = fnspecppk_lsl, 
  #              usl = fnspecppk_usl, 
  #              target = fnspecppk_nom)

  ppkout <- pcr(fndatppk$value, 
                "normal", 
                lsl = fnspecppk_lsl, 
                usl = fnspecppk_usl, 
                target = fnspecppk_nom)
  
  
  ppkout_list = list(ppkout)
  

  # ppkout <- ss.study.ca(xST = fndatppk$value, 
  #                       LSL = fnspecppk_lsl,
  #                       USL = fnspecppk_usl,
  #                       Target = fnspecppk_nom)
  
  # str(ppkout)
  
  return(ppkout)
  
}


# warnings()

# function 3 ------------------------------------------------------------------
# normal probability plot with confidence bands
qqplot_fn <- function(fndatnorm){
  
  require(qqplotr)
  
  # TEST
  # fndatnorm <- d_dat_nestj$data[[1]]
  
  
  fndatnorm <- fndatnorm %>% rename_at(2, ~"value")
  
  ggout <- fndatnorm %>% 
    ggplot(mapping = aes(sample = value)) + 
    stat_qq_band() +
    stat_qq_line() +
    stat_qq_point() +
    labs(x = "Theoretical Quantiles", y = "Sample Quantiles") + 
    ggtitle("Normal Probability Plot") + 
    NULL
  
  
  
  
  ggout <- ggout + 
    annotate("text", x = 0.9*max(x), y = .9*max(density(x)$y), label = xpars, hjust = 0)
  
  
  return(ggout)
  
}


  


# READ DATA ###################################################################

fname <- file.choose()

fname <- "074364 4-8-2020 TEST.xlsx"
fsheet <- "DATA SHT 1"



# data-------------------------------------------------------------------------
d_dat_wide0 <- read_xlsx(path = fname, sheet = fsheet, skip = 11) 


d_dat_wide <- d_dat_wide0 %>% 
  clean_names() %>% 
  mutate_if(is.logical, as.numeric) %>% 
  mutate_if(is.character, as.numeric)



# header info -----------------------------------------------------------------
x_job_name <- names(read_xlsx(path = fname, sheet = fsheet, range = "A3"))
x_part_num <- names(read_xlsx(path = fname, sheet = fsheet, range = "A4"))
x_inspector <- names(read_xlsx(path = fname, sheet = fsheet, range = "F3"))
x_insp_date <- names(read_xlsx(path = fname, sheet = fsheet, range = "F4"))



# spec info -------------------------------------------------------------------

d_spec_wide0 <- read_xlsx(path = fname, sheet = fsheet, 
                          col_names = FALSE, 
                          range = cell_rows(c(6:9)))

names(d_spec_wide0) <- c('spec_type', names(d_dat_wide)[2:dim(d_dat_wide)[2]])[1:dim(d_spec_wide0)[2]]


# drop "tolerance" row and convert 'feature' columns to numeric
d_spec_wide <- d_spec_wide0 %>% 
  filter(spec_type != "Tolerance") %>% 
  mutate_at(vars(starts_with("feature")), as.numeric)



# convert spec to long data for later joining 

# get rownames to use later as column names
x_rownames <- tolower(t(as.matrix(d_spec_wide[,1])))
x_colnames <- names(d_spec_wide)

# convert values to matrix for transpose
d_spec_tw <- as.data.frame(t(as.matrix(d_spec_wide[, c(-1)]))) 

# add column names
names(d_spec_tw) <- x_rownames

# convert dataframe rownames to column
d_spec_tw <- d_spec_tw %>% 
  tibble::rownames_to_column("feature")


d_spec_list <- d_spec_tw %>% 
  mutate(speclist = Map(c, nominal, lsl, usl))


d_spec_list$speclist[[1]]



# data wide to long -----------------------------------------------------------

colnames <- names(d_dat_wide[2:dim(d_dat_wide)[2]])

d_dat_long <- d_dat_wide %>% 
  pivot_longer(cols = starts_with("feature"), 
               names_to = "feature", 
               values_to = "value", 
               values_drop_na = TRUE) %>% 
  arrange(feature, specimen)






# # nested data =================================================================
# 
# d_dat_nest <- d_dat_long %>% 
#   group_by(feature) %>% 
#   nest()
# 
# d_dat_nest$data[1]
# 
# 
# # join data with spec info ----------------------------------------------------
# 
# d_dat_nestj <- d_dat_nest %>% 
#   left_join(d_spec_list, by = "feature") # %>% 
#   
# # d_dat_nestj$data[[1]]$value
# 
# 
# 
# # get process capability ======================================================

# # TEST function
# # fn_capability_plot(d_dat_nestj$data[[1]], d_dat_nestj$speclist[[1]])
# 
# 
# 
# d_dat_nest2 <- d_dat_nestj %>% 
#   mutate(ppkplot = map2(data, speclist, fn_capability_plot)) %>% 
#   mutate(ppkcalc = map2(data, speclist, fn_ppk)) %>% 
#   mutate(normplot = map(data, qqplot_fn))
# 
# 
# # model = map(data, lm_model)
# # verify
# # plot(d_dat_nest2$ppkplot[[3]])
# # plot(d_dat_nest2$normplot[[3]])


## new approach using output from loop ----------------------------------------

library(SixSigma)
help(package = "SixSigma")
?qualityTools
?qcc
help(package = 'ggQC')


ss.study.ca(xST = d_dat_nestj$data[[1]]$value, 
            xLT = d_dat_nestj$data[[1]], 
            LSL = d_dat_nestj$speclist[[1]][1],
            USL = d_dat_nestj$speclist[[1]][3],
            Target = d_dat_nestj$speclist[[1]][2], 
            f.sub = d_dat_nestj$feature[1])

 


# ANALYSIS ####################################################################





# OUTPUT ######################################################################
# save as image/object to Excel DIR in new sheet
# printable pdf
# archive in network folder




# SAVE DATA IN DB #############################################################





# END CODE ####################################################################