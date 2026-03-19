# ==========================================================
# Summaries (RoB + Yes/No) & Traffic plots (full and paged)
# One run generates all images at once
# - Only the summaries are combined; traffic is not combined
# - No plot titles anywhere
# ==========================================================

# ▶ Packages
pkgs <- c("readxl","dplyr","tidyr","stringr","ggplot2","scales",
          "forcats","rlang","ggnewscale")
to_install <- pkgs[!(pkgs %in% installed.packages()[, "Package"])]
if (length(to_install)) install.packages(to_install, dependencies = TRUE)
invisible(lapply(pkgs, library, character.only = TRUE))

# ▶ Inputs (adjust to your files/sheets)
xlsx_rob <- "risk_of_bias.xlsx"            # Low/Unclear/High
sheet_rob <- NULL                          # e.g. "RoB"

xlsx_yn  <- "additional_bias_items.xlsx"   # Yes/No
sheet_yn <- NULL                           # e.g. "YesNo"

# Export options
dpi              <- 600
width_summary_in <- 16    # width for summaries
width_grid_in    <- 18    # width for traffic (wider for labels)
row_height_in    <- 0.55  # summary bar thickness per domain row (inches)
tile_height_in   <- 0.28  # traffic tile height per study (inches)
gap_between_sum  <- 0.40  # white space between the two summaries when combined (inches)
page_block_size  <- 10    # studies per paged traffic image

# Output files
out_summary_rob <- "summary_rob.png"
out_summary_yn  <- "summary_yesno.png"
out_combined    <- "summary_rob_plus_yesno_same_area.png"

out_grid_rob    <- "traffic_rob.png"
out_grid_yn     <- "traffic_yesno.png"
# paged outputs: traffic_rob_page_XX.png and traffic_yesno_page_XX.png

# ▶ Palettes
pal_rob <- c("Low"="#80CBC4","Unclear"="#FFD54F","High"="#E57373")
pal_yn  <- c("No" ="#FFB74D","Yes"="#64B5F6")

# ▶ Normalizers
norm3 <- function(x){
  y <- tolower(trimws(as.character(x)))
  map <- c("low"="Low","l"="Low","probably low"="Low",
           "unclear"="Unclear","unknown"="Unclear","uncertain"="Unclear",
           "n/a"="Unclear","na"="Unclear","some concerns"="Unclear",
           "high"="High","h"="High","probably high"="High",
           "baixo"="Low","baixa"="Low","incerto"="Unclear","incerta"="Unclear",
           "alto"="High","alta"="High")
  out <- unname(map[y]); out[is.na(out) & !is.na(y)] <- NA_character_
  factor(out, levels = c("Low","Unclear","High"), ordered = TRUE)
}
norm_yn <- function(x){
  y <- tolower(trimws(as.character(x)))
  yes <- c("yes","y","true","t","1","sim","sí","si","verdadeiro","verdadero","present")
  no  <- c("no","n","false","f","0","não","nao","na","n/a","absent","not reported")
  out <- ifelse(y %in% yes, "Yes", ifelse(y %in% no, "No", NA_character_))
  factor(out, levels = c("No","Yes"), ordered = TRUE)
}

# ▶ Core builder: read, clean, normalize; return long + summary + orders
build_data <- function(xlsx_path, sheet_name, normalizer){
  dat <- readxl::read_excel(xlsx_path, sheet = sheet_name, .name_repair = "unique")
  names(dat) <- names(dat) |> stringr::str_squish()
  if (!"Study" %in% names(dat)) rlang::abort(paste0("'", xlsx_path, "' needs a 'Study' column."))
  
  # Clean
  overall_idx <- which(tolower(names(dat)) == "overall")
  if (length(overall_idx)) names(dat)[overall_idx] <- "Overall"
  dat <- dat |>
    mutate(
      Study = str_squish(str_replace_all(as.character(Study), "_", " ")),
      across(where(is.character), str_squish)
    ) |>
    filter(!is.na(Study) & Study != "")
  
  # Domain order
  domain_cols <- setdiff(names(dat), "Study")
  if (!length(domain_cols)) rlang::abort(paste0("No domain columns in '", xlsx_path, "'."))
  overall_pos <- which(tolower(domain_cols) == "overall")
  if (length(overall_pos) == 1) domain_cols <- c(domain_cols[-overall_pos], domain_cols[overall_pos])
  
  # Normalize judgments
  dat <- dat |> mutate(across(all_of(domain_cols), normalizer))
  
  # Long
  long <- dat |>
    pivot_longer(all_of(domain_cols), names_to="Domain_raw", values_to="Judgment") |>
    mutate(Domain = factor(gsub("\\.", " ", Domain_raw),
                           levels = gsub("\\.", " ", domain_cols)))
  
  # Study ordering (use Overall when present)
  if (length(overall_pos) == 1) {
    overall_display <- gsub("\\.", " ", domain_cols[length(domain_cols)])
    ov <- long |> filter(Domain == overall_display) |> mutate(score = as.integer(Judgment))
    # For RoB: High(3) worst on top; For Yes/No: No(1) on top
    # We'll decide outside by reading the levels of 'Judgment'
    if (all(levels(long$Judgment) %in% c("Low","Unclear","High"))) {
      study_order <- ov |> arrange(desc(score), Study) |> pull(Study) |> unique()
    } else {
      study_order <- ov |> arrange(score, Study) |> pull(Study) |> unique()
    }
  } else {
    study_order <- sort(unique(dat$Study))
  }
  study_order <- unique(c(study_order, setdiff(dat$Study, study_order)))
  long <- long |> mutate(Study = factor(Study, levels = rev(study_order)))
  
  # Summary df
  summary_df <- long |>
    filter(!is.na(Judgment)) |>
    count(Domain, Judgment, name = "n", .drop = FALSE) |>
    group_by(Domain) |>
    mutate(pct = if (sum(n) > 0) n/sum(n) else 0) |>
    ungroup()
  
  list(
    long = long,
    summary = summary_df,
    n_domains = nlevels(summary_df$Domain),
    n_studies = nlevels(long$Study)
  )
}

# ▶ Build RoB and Yes/No datasets
rob <- build_data(xlsx_rob, sheet_rob, norm3)
yn  <- build_data(xlsx_yn,  sheet_yn,  norm_yn)

# ---------- SUMMARY PLOTS (separate) ----------
make_summary_plot <- function(summary_df, pal){
  lv <- levels(droplevels(summary_df$Judgment))
  vals <- pal[names(pal) %in% lv]; vals <- vals[lv]
  ggplot(summary_df, aes(x=pct, y=Domain, fill=Judgment)) +
    geom_col(color="grey60", linewidth=0.2, width=0.90) +
    scale_x_continuous(labels=scales::percent_format(accuracy=1),
                       expand=expansion(mult=c(0,0.04))) +
    scale_fill_manual(values=vals, breaks=lv, labels=lv, name="Judgment") +
    labs(x=NULL, y=NULL) +
    theme_minimal(base_size=13) +
    theme(
      legend.position="right",
      panel.grid.major.y=element_blank(),
      panel.grid.minor  =element_blank()
    ) +
    geom_text(aes(label=ifelse(pct>=0.03, scales::percent(pct, accuracy=1), "")),
              position=position_stack(vjust=0.5), size=3)
}

p_summary_rob <- make_summary_plot(rob$summary, pal_rob)
p_summary_yn  <- make_summary_plot(yn$summary,  pal_yn)

h_summary_rob <- max(6, rob$n_domains * row_height_in)
h_summary_yn  <- max(6, yn$n_domains  * row_height_in)

ggsave(out_summary_rob, p_summary_rob, width=width_summary_in, height=h_summary_rob, units="in", dpi=dpi)
ggsave(out_summary_yn,  p_summary_yn,  width=width_summary_in, height=h_summary_yn,  units="in", dpi=dpi)

message("Saved: ", normalizePath(out_summary_rob))
message("Saved: ", normalizePath(out_summary_yn))

# ---------- COMBINED SUMMARY (same plotting area; same row thickness; NO titles/strips) ----------
sum_rob <- rob$summary |> mutate(Panel = factor("RoB", levels=c("RoB","Yes/No")))
sum_yn  <- yn$summary  |> mutate(Panel = factor("Yes/No", levels=c("RoB","Yes/No")))

p_combined <- ggplot() +
  facet_grid(Panel ~ ., scales="free_y", space="free_y", switch="y") +
  # RoB
  geom_col(data=sum_rob, aes(x=pct, y=Domain, fill=Judgment),
           color="grey60", linewidth=0.2, width=0.90) +
  scale_x_continuous(labels=scales::percent_format(accuracy=1),
                     expand=expansion(mult=c(0,0.04))) +
  scale_fill_manual(values=pal_rob, name="Judgment (RoB)", limits=c("Low","Unclear","High")) +
  geom_text(data=sum_rob,
            aes(x=pct, y=Domain, label=ifelse(pct>=0.03, scales::percent(pct, accuracy=1), "")),
            position=position_stack(vjust=0.5), size=3) +
  ggnewscale::new_scale_fill() +
  # Yes/No
  geom_col(data=sum_yn, aes(x=pct, y=Domain, fill=Judgment),
           color="grey60", linewidth=0.2, width=0.90) +
  scale_fill_manual(values=pal_yn, name="Judgment (Yes/No)", limits=c("No","Yes")) +
  geom_text(data=sum_yn,
            aes(x=pct, y=Domain, label=ifelse(pct>=0.03, scales::percent(pct, accuracy=1), "")),
            position=position_stack(vjust=0.5), size=3) +
  labs(x=NULL, y=NULL) +
  theme_minimal(base_size=13) +
  theme(
    legend.position = "right",
    panel.grid.major.y = element_blank(),
    panel.grid.minor   = element_blank(),
    strip.background   = element_blank(),
    strip.text.y.left  = element_blank(),   # ← no facet labels (no titles)
    panel.spacing.y    = unit(0.8, "lines")
  )

h_combined <- h_summary_rob + gap_between_sum + h_summary_yn
ggsave(out_combined, p_combined, width=width_summary_in, height=h_combined, units="in", dpi=dpi)
message("Saved: ", normalizePath(out_combined))

# ---------- TRAFFIC PLOTS (full) ----------
make_traffic_plot <- function(long_df, pal){
  lv <- levels(droplevels(long_df$Judgment))
  vals <- pal[names(pal) %in% lv]; vals <- vals[lv]
  ggplot(long_df, aes(x=Domain, y=Study, fill=Judgment)) +
    geom_tile(color="white", linewidth=0.7, width=0.995, height=0.995) +
    scale_fill_manual(values=vals, breaks=lv, labels=lv, name="Judgment") +
    labs(x=NULL, y=NULL) +
    theme_minimal(base_size=13) +
    theme(
      axis.text.x = element_text(angle=45, hjust=1, vjust=1),
      legend.position="right",
      panel.grid=element_blank()
    ) +
    coord_fixed(ratio=1)
}

p_grid_rob <- make_traffic_plot(rob$long, pal_rob)
p_grid_yn  <- make_traffic_plot(yn$long,  pal_yn)

h_grid_rob <- max(10, rob$n_studies * tile_height_in)
h_grid_yn  <- max(10, yn$n_studies  * tile_height_in)

ggsave(out_grid_rob, p_grid_rob, width=width_grid_in, height=h_grid_rob, units="in", dpi=dpi)
ggsave(out_grid_yn,  p_grid_yn,  width=width_grid_in, height=h_grid_yn,  units="in", dpi=dpi)

message("Saved: ", normalizePath(out_grid_rob))
message("Saved: ", normalizePath(out_grid_yn))

# ---------- TRAFFIC PLOTS (paged/blocks) ----------
save_paged_traffic <- function(long_df, pal, prefix){
  studies_all <- levels(long_df$Study)
  chunks <- split(studies_all, ceiling(seq_along(studies_all) / page_block_size))
  i <- 1L
  for (block in chunks) {
    long_i <- long_df |> dplyr::filter(Study %in% block) |> dplyr::mutate(Study = forcats::fct_drop(Study))
    p_i <- make_traffic_plot(long_i, pal)
    n_studies_i <- length(block)
    page_width  <- max(14, nlevels(long_i$Domain) * 0.7)
    page_height <- max(6,  n_studies_i * 0.35)
    f_i <- sprintf("%s_page_%02d.png", prefix, i)
    ggsave(f_i, p_i, width=page_width, height=page_height, units="in", dpi=dpi)
    message("Saved: ", normalizePath(f_i))
    i <- i + 1L
  }
}

save_paged_traffic(rob$long, pal_rob, "traffic_rob")
save_paged_traffic(yn$long,  pal_yn,  "traffic_yesno")

# ==========================================================
# Combined Summary (RoB + Yes/No) with customizations
# ==========================================================

# ... keep the same data building functions as before ...

# ▶ Custom relabeling function for Yes/No domains
relabel_yesno <- function(x) {
  x <- gsub("\\.", " ", x)
  case_when(
    grepl("^Other pre", x, ignore.case = TRUE) ~ "Preregistration",
    grepl("^Other impossibility", x, ignore.case = TRUE) ~ "Blinding impossible",
    TRUE ~ x
  )
}

# Apply relabeling
sum_rob <- rob$summary |> 
  mutate(Panel = factor("RoB", levels=c("RoB","Yes/No")),
         Domain = factor(as.character(Domain), 
                         levels = rev(levels(Domain))))   # invert order

sum_yn  <- yn$summary  |> 
  mutate(Panel = factor("Yes/No", levels=c("RoB","Yes/No")),
         Domain = relabel_yesno(as.character(Domain))) |> 
  mutate(Domain = factor(Domain, levels = rev(unique(Domain))))  # invert order

# ▶ Combined plot
p_combined <- ggplot() +
  facet_grid(Panel ~ ., scales="free_y", space="free_y", switch="y") +
  
  # --- RoB ---
  geom_col(data=sum_rob, aes(x=pct, y=Domain, fill=Judgment),
           color="grey60", linewidth=0.2, width=0.8) +   # thinner bars
  scale_x_continuous(labels=scales::percent_format(accuracy=1),
                     expand=expansion(mult=c(0,0.04))) +
  scale_fill_manual(values=pal_rob,
                    breaks=c("Low","Unclear","High"),
                    labels=c("Low","Unclear","High")) +
  geom_text(data=sum_rob,
            aes(x=pct, y=Domain, 
                label=ifelse(pct>=0.03, scales::percent(pct, accuracy=1), "")),
            position=position_stack(vjust=0.5), size=3.2) +
  
  # --- Yes/No ---
  ggnewscale::new_scale_fill() +
  geom_col(data=sum_yn, aes(x=pct, y=Domain, fill=Judgment),
           color="grey60", linewidth=0.2, width=0.8) +
  scale_fill_manual(values=pal_yn,
                    breaks=c("Yes","No"), 
                    labels=c("Yes","No")) +
  geom_text(data=sum_yn,
            aes(x=pct, y=Domain, 
                label=ifelse(pct>=0.03, scales::percent(pct, accuracy=1), "")),
            position=position_stack(vjust=0.5), size=3.2) +
  
  # Labels and theme
  labs(x=NULL, y=NULL) +
  theme_minimal(base_size=15) +   # bigger base font
  theme(
    legend.position="right",
    panel.grid.major.y=element_blank(),
    panel.grid.minor=element_blank(),
    strip.background=element_blank(),
    strip.text.y.left=element_blank(),
    axis.text.y=element_text(size=13),  # larger domain font
    panel.spacing.y=unit(2, "lines")    # more space between RoB and Yes/No
  )

# Adjust height to preserve bar thickness
h_combined <- (rob$n_domains + yn$n_domains) * row_height_in + 3
ggsave(out_combined, p_combined, width=width_summary_in, height=h_combined, units="in", dpi=dpi)
message("Saved: ", normalizePath(out_combined))

library(dplyr)
library(ggplot2)
library(ggnewscale)
library(scales)

# ▶ Paletas
pal_rob <- c("Low"="#80CBC4","Unclear"="#FFD54F","High"="#E57373")
pal_yn  <- c("Yes"="#64B5F6","No"="#FFB74D")

# ▶ Relabels por painel
relabel_rob <- function(x) {
  x <- gsub("\\.", " ", x)
  dplyr::case_when(
    grepl("^Other pseudorep", x, ignore.case = TRUE)     ~ "Pseudoreplication",
    grepl("^Other procedural eq", x, ignore.case = TRUE) ~ "Procedural equivalence",
    TRUE ~ x
  )
}

relabel_yesno <- function(x) {
  x <- gsub("\\.", " ", x)
  dplyr::case_when(
    grepl("^Other pre", x, ignore.case = TRUE)           ~ "Preregistration",
    grepl("^Other impossibility", x, ignore.case = TRUE) ~ "Blinding impossible",
    TRUE ~ x
  )
}

# RoB
sum_rob <- rob$summary %>%
  mutate(
    Panel    = factor("RoB", levels = c("RoB","Yes/No")),
    Domain   = relabel_rob(as.character(Domain)),                   # <- AQUI
    Domain   = factor(Domain, levels = rev(unique(Domain))),
    Judgment = factor(Judgment, levels = c("Low","Unclear","High"))
  )

# Yes/No
sum_yn <- yn$summary %>%
  mutate(
    Panel    = factor("Yes/No", levels = c("RoB","Yes/No")),
    Domain   = relabel_yesno(as.character(Domain)),                 # <- SÓ estes dois
    Domain   = factor(Domain, levels = rev(unique(Domain))),
    Judgment = factor(Judgment, levels = c("Yes","No"))
  )


# ----------------- PLOT -----------------
p_combined <- ggplot() +
  facet_grid(Panel ~ ., scales = "free_y", space = "free_y", switch = "y") +
  
  # --- RoB ---
  geom_col(
    data = sum_rob,
    aes(x = pct, y = Domain, fill = Judgment),
    color = "grey60", linewidth = 0.2, width = 0.70,
    position = position_stack(reverse = TRUE)
  ) +
  scale_x_continuous(labels = percent_format(accuracy = 1),
                     expand = expansion(mult = c(0, 0.05))) +
  scale_fill_manual(values = pal_rob, breaks = c("Low","Unclear","High"), name = "Judgment") +
  geom_text(
    data = sum_rob,
    aes(x = pct, y = Domain, label = ifelse(pct >= 0.03, percent(pct, accuracy = 1), "")),
    position = position_stack(vjust = 0.5, reverse = FALSE),
    size = 4.0
  ) +
  
  # --- Yes/No ---
  ggnewscale::new_scale_fill() +
  geom_col(
    data = sum_yn,
    aes(x = pct, y = Domain, fill = Judgment),
    color = "grey60", linewidth = 0.2, width = 0.70,
    position = position_stack(reverse = TRUE)
  ) +
  scale_fill_manual(values = pal_yn, breaks = c("Yes","No"), name = NULL) +
  geom_text(
    data = sum_yn,
    aes(x = pct, y = Domain, group = Judgment,
        label = ifelse(pct >= 0.03, scales::percent(pct, accuracy = 1), "")),
    position = position_stack(vjust = 0.5, reverse = TRUE),  # ← inverter aqui
    size = 4.0
  ) +
  
  labs(x = NULL, y = NULL) +
  theme_minimal(base_size = 16) +
  theme(
    plot.background  = element_rect(fill = "white", color = NA),
    panel.background = element_rect(fill = "white", color = NA),
    
    strip.background = element_blank(),
    strip.text.y.left = element_blank(),
    
    panel.grid.major.y = element_blank(),
    panel.grid.minor = element_blank(),
    
    axis.text.y = element_text(size = 17),
    
    legend.position = "right",
    panel.spacing.y = unit(2, "lines")
  )

h_combined <- (rob$n_domains + yn$n_domains) * row_height_in + 3
ggsave(out_combined, p_combined,
       width = width_summary_in, height = h_combined,
       units = "in", dpi = dpi)
message("Saved: ", normalizePath(out_combined))
