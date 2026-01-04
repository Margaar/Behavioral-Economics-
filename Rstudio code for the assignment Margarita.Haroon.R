
# ---- packages
pkgs <- c("readxl","dplyr","ggplot2","janitor","stringr","broom","car","sandwich","lmtest","tibble")
to_install <- pkgs[!pkgs %in% installed.packages()[,"Package"]]
if(length(to_install)>0) install.packages(to_install)

library(readxl); library(dplyr); library(ggplot2); library(janitor)
library(stringr); library(broom); library(car); library(sandwich)
library(lmtest); library(tibble)


file_path <- "C:/Users/mg/Desktop/Rstudio/The Mystery Box Effect in Consumer Behavior (Responses) (1).xlsx"

out_dir <- "outputs_mystery_box"
if(!dir.exists(out_dir)) dir.create(out_dir)


df_raw <- read_excel(file_path, sheet = "Form Responses 1") |> clean_names()
cat("Columns:\n"); print(names(df_raw))

#scenario column
col_scn <- "you_are_randomly_assigned_to_one_of_two_scenarios"


col_wtp_known   <- names(df_raw)[3]
col_wtp_mystery <- names(df_raw)[4]

# ---- build unified dataset 
df <- df_raw |>
  mutate(
    condition = ifelse(.data[[col_scn]] == "Scenario A", "Known product", "Mystery box"),
    wtp = ifelse(condition=="Known product", .data[[col_wtp_known]], .data[[col_wtp_mystery]]),
    wtp = suppressWarnings(as.numeric(str_replace_all(as.character(wtp), ",","."))),
    age = suppressWarnings(as.numeric(your_age)),
    curiosity = suppressWarnings(as.numeric(on_a_scale_of_1_7_how_curious_are_you_generally_1_not_curious_at_all_7_extremely_curious)),
    prior = as.factor(have_you_ever_purchased_a_mystery_box_surprise_item_or_blind_box_product),
    seriousness = as.factor(how_seriously_did_you_consider_your_bid),
    gender = as.factor(gender),
    frequency = as.factor(frequency)
  ) |>
  filter(!is.na(wtp) & wtp >= 0)

write.csv(df, file.path(out_dir,"cleaned_data.csv"), row.names=FALSE)
cat("\nN by condition:\n"); print(table(df$condition))

known_df   <- df |> filter(condition=="Known product")
mystery_df <- df |> filter(condition=="Mystery box")

#descriptive stats
desc <- df |> group_by(condition) |>
  summarise(
    N=n(),
    mean=mean(wtp), sd=sd(wtp), median=median(wtp),
    min=min(wtp), max=max(wtp),
    q25=quantile(wtp,0.25), q75=quantile(wtp,0.75),
    .groups="drop"
  )
write.csv(desc, file.path(out_dir,"desc_wtp.csv"), row.names=FALSE)

#main tests (Welch + Mann-Whitney + Levene + Cohen's d)
tt <- t.test(wtp ~ condition, data=df, var.equal=FALSE) # Welch
mw <- wilcox.test(wtp ~ condition, data=df, exact=FALSE)
lev <- leveneTest(wtp ~ condition, data=df, center=median)

cohen_d <- function(x,y){
  nx <- length(x); ny <- length(y)
  sx <- sd(x); sy <- sd(y)
  sp <- sqrt(((nx-1)*sx^2+(ny-1)*sy^2)/(nx+ny-2))
  (mean(x)-mean(y))/sp
}
d <- cohen_d(mystery_df$wtp, known_df$wtp)

main_tests <- data.frame(
  mean_known=mean(known_df$wtp),
  mean_mystery=mean(mystery_df$wtp),
  diff_mean=mean(mystery_df$wtp)-mean(known_df$wtp),
  welch_t=unname(tt$statistic),
  welch_df=unname(tt$parameter),
  welch_p=tt$p.value,
  ci_low=tt$conf.int[1],
  ci_high=tt$conf.int[2],
  mann_whitney_W=unname(mw$statistic),
  mann_whitney_p=mw$p.value,
  cohen_d=d,
  levene_p=broom::tidy(lev)$p.value[1]
)
write.csv(main_tests, file.path(out_dir,"main_tests.csv"), row.names=FALSE)

#Known block extra analysis: effect of prior experience
known_prior_test <- t.test(wtp ~ prior, data=known_df, var.equal=FALSE)
known_prior_desc <- known_df |> group_by(prior) |> summarise(N=n(), mean=mean(wtp), sd=sd(wtp), .groups="drop")
write.csv(known_prior_desc, file.path(out_dir,"known_prior_desc.csv"), row.names=FALSE)
write.csv(broom::tidy(known_prior_test), file.path(out_dir,"known_prior_test.csv"), row.names=FALSE)

#plots
p1 <- ggplot(df, aes(condition, wtp)) +
  geom_jitter(width=0.15, alpha=0.6) +
  stat_summary(fun=mean, geom="point", shape=18, size=4) +
  labs(title="WTP by condition", x="", y="WTP (EUR)")
ggsave(file.path(out_dir,"wtp_points_means.png"), p1, width=7, height=5, dpi=300)

p2 <- ggplot(df, aes(condition, wtp)) +
  geom_boxplot(outlier.alpha=0.3) +
  labs(title="WTP distribution by condition", x="", y="WTP (EUR)")
ggsave(file.path(out_dir,"wtp_boxplot.png"), p2, width=7, height=5, dpi=300)

#regressions with robust SE (FIXED robust_table)
df <- df |> mutate(mystery = ifelse(condition=="Mystery box",1,0))

robust_table <- function(model){
  V  <- sandwich::vcovHC(model, type = "HC1")
  ct <- lmtest::coeftest(model, vcov. = V)   
  
  m <- as.matrix(ct)  # matrix
  
  # 1scenario
  out <- data.frame(
    term      = rownames(m),
    estimate  = m[, 1],
    robust_se = m[, 2],
    statistic = m[, 3],
    stringsAsFactors = FALSE
  )
  
  if (ncol(m) >= 4) {
    out$p_value <- m[, 4]
  } else {
    # Chechking p-value if no use t tesr 
    df_res <- df.residual(model)
    out$p_value <- 2 * pt(abs(out$statistic), df = df_res, lower.tail = FALSE)
  }
  
  out
}


# FINAL: auto summary conclusions (printed + saved)
summary_lines <- c(
  "=== SUMMARY CONCLUSIONS (auto-generated) ===",
  paste0("Total N = ", nrow(df), " | Known N = ", nrow(known_df), " | Mystery N = ", nrow(mystery_df)),
  paste0("Known: mean WTP = ", round(mean(known_df$wtp),2), " (SD ", round(sd(known_df$wtp),2), "), median = ", round(median(known_df$wtp),2),
         ", range [", min(known_df$wtp), ", ", max(known_df$wtp), "]"),
  paste0("Mystery: mean WTP = ", round(mean(mystery_df$wtp),2), " (SD ", round(sd(mystery_df$wtp),2), "), median = ", round(median(mystery_df$wtp),2),
         ", range [", min(mystery_df$wtp), ", ", max(mystery_df$wtp), "]"),
  paste0("Mean difference (Mystery - Known) = ", round(mean(mystery_df$wtp)-mean(known_df$wtp),2), " EUR"),
  paste0("Welch t-test: t = ", round(unname(tt$statistic),3), ", df = ", round(unname(tt$parameter),2), ", p = ", signif(tt$p.value,3),
         ", CI [", round(tt$conf.int[1],2), ", ", round(tt$conf.int[2],2), "]"),
  paste0("Mann–Whitney: W = ", unname(mw$statistic), ", p = ", signif(mw$p.value,3)),
  paste0("Effect size: Cohen's d = ", round(d,3)),
  paste0("Variance (Levene/Brown–Forsythe): p = ", signif(broom::tidy(lev)$p.value[1],3)),
  "Known block (prior experience):",
  paste0("  Means by prior: ", paste0(known_prior_desc$prior, "=", round(known_prior_desc$mean,2), collapse=" | ")),
  paste0("  Welch test prior effect: t = ", round(known_prior_test$statistic,3), ", df = ", round(known_prior_test$parameter,2),
         ", p = ", signif(known_prior_test$p.value,3)),
  "Regression notes:",
  paste0("  In m0 (wtp~mystery), coefficient on mystery is in reg_m0_robust.csv"),
  paste0("  In m1 (with controls), coefficient on mystery is in reg_m1_robust.csv"),
  paste0("Outputs saved in folder: ", out_dir)
)

cat(paste(summary_lines, collapse="\n"), "\n")
writeLines(summary_lines, file.path(out_dir,"conclusions_for_paper.txt"))

cat("\nDONE. Outputs saved to:", out_dir, "\n")


# ---- Mystery block extra analysis: effect of prior experience (same as Known)
mystery_prior_test <- t.test(wtp ~ prior, data=mystery_df, var.equal=FALSE)
mystery_prior_desc <- mystery_df |> group_by(prior) |> summarise(N=n(), mean=mean(wtp), sd=sd(wtp), .groups="drop")

write.csv(mystery_prior_desc, file.path(out_dir,"mystery_prior_desc.csv"), row.names=FALSE)
write.csv(broom::tidy(mystery_prior_test), file.path(out_dir,"mystery_prior_test.csv"), row.names=FALSE)


#
# TWO CLEAN COMPARISON TABLES

# packages for tables
if(!"knitr" %in% installed.packages()[,"Package"]) install.packages("knitr")
if(!"dplyr" %in% installed.packages()[,"Package"]) install.packages("dplyr")
library(knitr)
library(dplyr)

# helper formatters
fmt_mean_sd <- function(x, digits=2){
  x <- x[!is.na(x)]
  if(length(x)==0) return("NA")
  paste0(round(mean(x),digits), " (", round(sd(x),digits), ")")
}
fmt_median_iqr <- function(x, digits=2){
  x <- x[!is.na(x)]
  if(length(x)==0) return("NA")
  q <- quantile(x, c(0.25,0.75))
  paste0(round(median(x),digits), " [", round(q[1],digits), "; ", round(q[2],digits), "]")
}
fmt_npct <- function(n, denom, digits=1){
  if(is.na(n) || is.na(denom) || denom==0) return(paste0(0, " (0\\%)"))
  paste0(n, " (", round(100*n/denom, digits), "\\%)")
}
fmt_p <- function(p){
  if(is.na(p)) return("NA")
  if(p < 0.001) return("<0.001")
  sprintf("%.3f", p)
}

nK <- nrow(known_df)
nM <- nrow(mystery_df)

# =========================
# TABLE 1 (NUMERIC): WTP + age + curiosity
# =========================
# tests
p_wtp_welch <- t.test(wtp ~ condition, data=df, var.equal=FALSE)$p.value
p_wtp_wilc  <- wilcox.test(wtp ~ condition, data=df, exact=FALSE)$p.value

p_age <- tryCatch(t.test(age ~ condition, data=df, var.equal=FALSE)$p.value, error=function(e) NA)
p_cur <- tryCatch(t.test(curiosity ~ condition, data=df, var.equal=FALSE)$p.value, error=function(e) NA)

diff_mean_wtp <- mean(mystery_df$wtp, na.rm=TRUE) - mean(known_df$wtp, na.rm=TRUE)

table_numeric <- data.frame(
  Variable = c("N",
               "WTP mean (SD)",
               "WTP median [IQR]",
               "WTP min--max",
               "Age mean (SD)",
               "Curiosity mean (SD)"),
  Known = c(nK,
            fmt_mean_sd(known_df$wtp),
            fmt_median_iqr(known_df$wtp),
            paste0(min(known_df$wtp,na.rm=TRUE), "--", max(known_df$wtp,na.rm=TRUE)),
            fmt_mean_sd(known_df$age),
            fmt_mean_sd(known_df$curiosity)),
  Mystery = c(nM,
              fmt_mean_sd(mystery_df$wtp),
              fmt_median_iqr(mystery_df$wtp),
              paste0(min(mystery_df$wtp,na.rm=TRUE), "--", max(mystery_df$wtp,na.rm=TRUE)),
              fmt_mean_sd(mystery_df$age),
              fmt_mean_sd(mystery_df$curiosity)),
  Difference = c("",
                 round(diff_mean_wtp,2),
                 "",
                 "",
                 "",
                 ""),
  `p-value` = c("",
                fmt_p(p_wtp_welch),
                fmt_p(p_wtp_wilc),
                "",
                fmt_p(p_age),
                fmt_p(p_cur)),
  stringsAsFactors = FALSE
)

latex_numeric <- kable(
  table_numeric,
  format="latex", booktabs=TRUE, escape=FALSE,
  caption="Comparison of numeric variables across conditions",
  label="tab:numeric_compare"
)

writeLines(latex_numeric, file.path(out_dir, "table_numeric.tex"))
cat("\nSaved:", file.path(out_dir, "table_numeric.tex"), "\n")


# =========================
# TABLE 2 (CATEGORICAL): short binary shares
# =========================
# Create SHORT indicators (to keep table readable)
df_short <- df |>
  mutate(
    female = ifelse(gender == "Female", 1, 0),
    prior_yes = ifelse(prior %in% c("Yes","yes"), 1, 0),
    serious_high = ifelse(grepl("Very|Serious", as.character(seriousness)), 1, 0),
    freq_high = ifelse(grepl("Often|Very", as.character(frequency)), 1, 0)
  )

K <- df_short |> filter(condition=="Known product")
M <- df_short |> filter(condition=="Mystery box")

# proportions + tests (use Fisher for safety)
tab_female <- matrix(c(sum(K$female,na.rm=TRUE), nrow(K)-sum(K$female,na.rm=TRUE),
                       sum(M$female,na.rm=TRUE), nrow(M)-sum(M$female,na.rm=TRUE)), nrow=2, byrow=TRUE)
p_female <- fisher.test(tab_female)$p.value

tab_prior <- matrix(c(sum(K$prior_yes,na.rm=TRUE), nrow(K)-sum(K$prior_yes,na.rm=TRUE),
                      sum(M$prior_yes,na.rm=TRUE), nrow(M)-sum(M$prior_yes,na.rm=TRUE)), nrow=2, byrow=TRUE)
p_prior <- fisher.test(tab_prior)$p.value

tab_ser <- matrix(c(sum(K$serious_high,na.rm=TRUE), nrow(K)-sum(K$serious_high,na.rm=TRUE),
                    sum(M$serious_high,na.rm=TRUE), nrow(M)-sum(M$serious_high,na.rm=TRUE)), nrow=2, byrow=TRUE)
p_ser <- fisher.test(tab_ser)$p.value

tab_freq <- matrix(c(sum(K$freq_high,na.rm=TRUE), nrow(K)-sum(K$freq_high,na.rm=TRUE),
                     sum(M$freq_high,na.rm=TRUE), nrow(M)-sum(M$freq_high,na.rm=TRUE)), nrow=2, byrow=TRUE)
p_freq <- fisher.test(tab_freq)$p.value

table_cat <- data.frame(
  Variable = c("Female",
               "Prior mystery purchase = Yes",
               "High seriousness (Serious or Very serious)",
               "High chocolate purchase frequency (Often/Very often)"),
  Known = c(fmt_npct(sum(K$female,na.rm=TRUE), nrow(K)),
            fmt_npct(sum(K$prior_yes,na.rm=TRUE), nrow(K)),
            fmt_npct(sum(K$serious_high,na.rm=TRUE), nrow(K)),
            fmt_npct(sum(K$freq_high,na.rm=TRUE), nrow(K))),
  Mystery = c(fmt_npct(sum(M$female,na.rm=TRUE), nrow(M)),
              fmt_npct(sum(M$prior_yes,na.rm=TRUE), nrow(M)),
              fmt_npct(sum(M$serious_high,na.rm=TRUE), nrow(M)),
              fmt_npct(sum(M$freq_high,na.rm=TRUE), nrow(M))),
  `p-value` = c(fmt_p(p_female), fmt_p(p_prior), fmt_p(p_ser), fmt_p(p_freq)),
  stringsAsFactors = FALSE
)

latex_cat <- kable(
  table_cat,
  format="latex", booktabs=TRUE, escape=FALSE,
  caption="Comparison of categorical variables across conditions (shares)",
  label="tab:cat_compare"
)

writeLines(latex_cat, file.path(out_dir, "table_categorical.tex"))
cat("Saved:", file.path(out_dir, "table_categorical.tex"), "\n")
