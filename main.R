# 导入必要的包
library(readxl)      # 读取Excel文件
library(dplyr)       # 数据处理
library(writexl)     # 写入Excel文件
library(tibble)      # 用于add_column功能

# 处理年代数据并返回结果
process_age_data <- function(age, df, version, iugs_col, show_cols) {
  cat(sprintf("\n--- 正在处理年代: %s ---\n", age))
  
  # 添加调试信息 - 显示数据类型信息
  cat(sprintf("输入年代类型: %s, 值: %s\n", class(age), age))
  
  # 显示一些数据范围信息，帮助诊断问题
  tryCatch({
    version_vals <- df[[version]]
    version_numeric <- suppressWarnings(as.numeric(version_vals))
    valid_version_vals <- version_numeric[!is.na(version_numeric)]
    
    if(length(valid_version_vals) > 0) {
      cat(sprintf("版本列(%s)数值范围: 最小值=%s, 最大值=%s\n", 
                 version, min(valid_version_vals), max(valid_version_vals)))
    } else {
      cat(sprintf("警告: 版本列(%s)不包含有效数值\n", version))
    }
    
    iugs_vals <- df[[iugs_col]]
    iugs_numeric <- sapply(iugs_vals, function(x) suppressWarnings(as.numeric(as.character(x))))
    valid_iugs_vals <- iugs_numeric[!is.na(iugs_numeric)]
    
    if(length(valid_iugs_vals) > 0) {
      cat(sprintf("IUGS列(%s)数值范围: 最小值=%s, 最大值=%s\n", 
                 iugs_col, min(valid_iugs_vals), max(valid_iugs_vals)))
    } else {
      cat(sprintf("警告: IUGS列(%s)不包含有效数值\n", iugs_col))
    }
  }, error = function(e) {
    cat(sprintf("获取数据范围信息时出错: %s\n", e$message))
  })
  
  # 智能匹配年代
  match_age_func <- function(val) {
    if (is.na(val)) {
      return(FALSE)
    }
    
    # 确保处理字符串值
    val_str <- as.character(val)
    age_str <- as.character(age)
    
    tryCatch({
      # 首先进行字符串比较
      if (trimws(val_str) == trimws(age_str)) {
        return(TRUE)
      }
      
      # 尝试进行数值比较
      val_num <- suppressWarnings(as.numeric(val_str))
      age_num <- suppressWarnings(as.numeric(age_str))
      
      if (!is.na(val_num) && !is.na(age_num) && val_num == age_num) {
        return(TRUE)
      }
      
      return(FALSE)
    }, error = function(e) {
      # 如果无法转换为数值，只进行字符串比较
      return(trimws(val_str) == trimws(age_str))
    })
  }
  
  # 找到匹配的行
  match_idx <- NULL
  for (i in 1:nrow(df)) {
    if (match_age_func(df[[version]][i])) {
      match_idx <- i
      break
    }
  }
  
  if (!is.null(match_idx)) {
    # 有完全匹配
    row <- df[match_idx, ]
    labels <- c("宇", "界", "系", "统", "阶")
    info <- character(0)
    
    for (i in 1:length(labels)) {
      col_label <- labels[i]
      if (col_label %in% names(df)) {
        info <- c(info, paste0(row[[col_label]], labels[i]))
      }
    }
    
    iugs_value <- row[[iugs_col]]
    cat(sprintf("直接匹配: %s %s\n", paste(info, collapse = " "), iugs_value))
    
    result_df <- df[match_idx, show_cols, drop = FALSE]
    return(result_df)
  } else {
    # 尝试插值
    tryCatch({
      age_num <- as.numeric(age)
    }, error = function(e) {
      cat(sprintf('年代"%s"无法识别为数字，无法插值。\n', age))
      return(NULL)
    })
    
    # 数值转换成功
    age_num <- as.numeric(age)
    
    # 过滤掉非数字值
    version_numeric <- suppressWarnings(as.numeric(df[[version]]))
    valid_idx <- !is.na(version_numeric)
    version_numeric_filtered <- version_numeric[valid_idx]
    df_filtered <- df[valid_idx, ]
    
    if (length(version_numeric_filtered) == 0) {
      cat(sprintf('%s列没有可用于插值的数字数据\n', version))
      return(NULL)
    }
    
    # 找上下边界
    lower_idx_values <- which(version_numeric_filtered <= age_num)
    if (length(lower_idx_values) == 0) {
      cat(sprintf('无法找到下边界进行插值，输入年代值(%s)可能小于数据集最小值\n', age_num))
      # 尝试使用最小值作为下边界
      lower_idx <- which.min(version_numeric_filtered)
      lower_idx_original <- which(valid_idx)[lower_idx]
      cat(sprintf('尝试使用最小值 %s 作为下边界\n', version_numeric_filtered[lower_idx]))
    } else {
      lower_idx <- lower_idx_values[which.max(version_numeric_filtered[lower_idx_values])]
      lower_idx_original <- which(valid_idx)[lower_idx]
    }
    
    upper_idx_values <- which(version_numeric_filtered >= age_num)
    if (length(upper_idx_values) == 0) {
      cat(sprintf('无法找到上边界进行插值，输入年代值(%s)可能大于数据集最大值\n', age_num))
      # 尝试使用最大值作为上边界
      upper_idx <- which.max(version_numeric_filtered)
      upper_idx_original <- which(valid_idx)[upper_idx]
      cat(sprintf('尝试使用最大值 %s 作为上边界\n', version_numeric_filtered[upper_idx]))
    } else {
      upper_idx <- upper_idx_values[which.min(version_numeric_filtered[upper_idx_values])]
      upper_idx_original <- which(valid_idx)[upper_idx]
    }
    
    # 添加调试信息
    cat(sprintf("选择的下边界值: %s, 上边界值: %s\n", 
               version_numeric_filtered[lower_idx], 
               version_numeric_filtered[upper_idx]))
    
    if (lower_idx == upper_idx) {
      if (df_filtered[[version]][lower_idx] == age_num) {
        # 精确匹配
        match_row_df <- df_filtered[lower_idx, ]
        labels <- c("宇", "界", "系", "统", "阶")
        info <- character(0)
        
        for (i in 1:length(labels)) {
          col_label <- labels[i]
          if (col_label %in% names(df)) {
            info <- c(info, paste0(match_row_df[[col_label]], labels[i]))
          }
        }
        
        iugs_value <- match_row_df[[iugs_col]]
        cat(sprintf("精确数值匹配: %s %s\n", paste(info, collapse = " "), iugs_value))
        
        result_df <- df[lower_idx_original, show_cols, drop = FALSE]
        result_df[[version]] <- age_num
        return(result_df)
      } else {
        cat(sprintf('无法找到合适的上下界进行插值，输入值 %s 可能超出 %s 列的数值范围。\n', age_num, version))
        return(NULL)
      }
    }
    
    # 进行插值
    lower_row <- df_filtered[lower_idx, ]
    upper_row <- df_filtered[upper_idx, ]
    lower_x <- as.numeric(lower_row[[version]])
    upper_x <- as.numeric(upper_row[[version]])
    
    # 显示IUGS值的调试信息
    cat(sprintf("下边界IUGS原始值: %s (类型: %s), 上边界IUGS原始值: %s (类型: %s)\n", 
               lower_row[[iugs_col]], class(lower_row[[iugs_col]]),
               upper_row[[iugs_col]], class(upper_row[[iugs_col]])))
    
    # 尝试将IUGS值转换为数字，处理可能是日期的情况
    safe_numeric_convert <- function(val) {
      # 首先检查输入值是否为NULL或长度为0
      if (is.null(val) || length(val) == 0) {
        cat(sprintf("警告: 输入值为NULL或长度为0\n"))
        return(NA)
      }
      
      # 然后检查输入值是否为NA
      if (is.na(val)) {
        cat(sprintf("警告: 输入值为NA\n"))
        return(NA)
      }
      
      # 尝试直接转换为数值
      result <- suppressWarnings(as.numeric(val))
      if (is.na(result)) {
        cat(sprintf("无法直接转换为数值: %s (类型: %s)\n", val, class(val)))
        
        # 如果是日期类型，尝试转换为数值
        if (inherits(val, "Date") || inherits(val, "POSIXt")) {
          cat(sprintf("检测到日期类型，尝试转换\n"))
          date_num <- as.numeric(as.POSIXct(val))
          return(date_num)
        }
        
        # 如果是字符串，尝试先转日期再转数值
        tryCatch({
          cat(sprintf("尝试将字符串转换为日期: %s\n", val))
          date_val <- as.Date(val)
          if (!is.na(date_val)) {
            date_num <- as.numeric(date_val)
            cat(sprintf("转换成功: %s -> %s\n", val, date_num))
            return(date_num)
          }
        }, error = function(e) {
          cat(sprintf("转换为日期失败: %s\n", e$message))
        })
        
        # 特殊情况: 尝试从字符串中提取数字
        tryCatch({
          num_extract <- as.numeric(gsub("[^0-9.]", "", as.character(val)))
          if (!is.na(num_extract)) {
            cat(sprintf("从字符串中提取数字: %s -> %s\n", val, num_extract))
            return(num_extract)
          }
        }, error = function(e) {
          cat(sprintf("提取数字失败: %s\n", e$message))
        })
      }
      return(result)
    }
    
    lower_y_val <- safe_numeric_convert(lower_row[[iugs_col]])
    upper_y_val <- safe_numeric_convert(upper_row[[iugs_col]])
    
    cat(sprintf("转换后的下边界IUGS值: %s, 上边界IUGS值: %s\n", lower_y_val, upper_y_val))
    
    # 修改NA值处理的逻辑 - 不直接返回NULL，而是尝试查找最接近的值
    if (is.na(lower_y_val) || is.na(upper_y_val)) {
      cat('上下界的IUGS 2023.9数据缺失或非数值，尝试查找最接近的值...\n')
      
      # 安全地转换IUGS列中的所有值为数值
      iugs_numeric <- sapply(df[[iugs_col]], safe_numeric_convert)
      valid_iugs_idx <- !is.na(iugs_numeric)
      
      if (sum(valid_iugs_idx) == 0) {
        cat('IUGS 2023.9列中没有有效的数值数据，无法处理。\n')
        return(NULL)
      }
      
      iugs_numeric_filtered <- iugs_numeric[valid_iugs_idx]
      df_iugs_filtered <- df[valid_iugs_idx, ]
      
      # 计算年代与该表中版本列每个值的距离
      version_numeric <- suppressWarnings(as.numeric(df_iugs_filtered[[version]]))
      valid_version_idx <- !is.na(version_numeric)
      
      if (sum(valid_version_idx) == 0) {
        cat('在IUGS有效数据的行中，没有有效的版本数值数据，无法处理。\n')
        return(NULL)
      }
      
      # 获取具有有效版本和IUGS值的行
      valid_rows <- df_iugs_filtered[valid_version_idx, ]
      valid_version_values <- version_numeric[valid_version_idx]
      
      # 计算与输入年代的距离
      age_num <- as.numeric(age)
      distances <- abs(valid_version_values - age_num)
      
      # 找到距离最小的行索引
      closest_idx <- which.min(distances)
      if (length(closest_idx) == 0) {
        cat('无法找到与输入年代最接近的值。\n')
        return(NULL)
      }
      
      closest_row_idx <- rownames(valid_rows)[closest_idx]
      closest_row <- valid_rows[closest_idx, ]
      closest_iugs_value <- closest_row[[iugs_col]]
      
      cat(sprintf("找到最接近年代 %s 的行，IUGS值为: %s\n", age, closest_iugs_value))
      
      # 创建结果DataFrame
      result_df <- df[as.numeric(closest_row_idx), show_cols, drop = FALSE]
      result_df[[version]] <- age_num
      
      # 显示信息
      labels <- c("宇", "界", "系", "统", "阶")
      info <- character(0)
      for (i in 1:length(labels)) {
        col_label <- labels[i]
        if (col_label %in% names(df)) {
          info <- c(info, paste0(closest_row[[col_label]], labels[i]))
        }
      }
      
      cat(sprintf("使用最接近的匹配: %s IUGS值:%s (输入年代:%s)\n", 
                 paste(info, collapse = " "), closest_iugs_value, age_num))
      
      return(result_df)
    }
    
    lower_y <- lower_y_val
    upper_y <- upper_y_val
    ratio <- ifelse(upper_x != lower_x, (age_num - lower_x) / (upper_x - lower_x), 0)
    interp_value <- lower_y + (upper_y - lower_y) * ratio
    
    cat(sprintf("插值计算: 下界=%s, 上界=%s, 比例=%s, 结果=%s\n", 
               lower_y, upper_y, ratio, interp_value))
    
    # 查找上界
    # 安全地转换IUGS列中的所有值为数值
    iugs_numeric <- sapply(df[[iugs_col]], safe_numeric_convert)
    valid_iugs_idx <- !is.na(iugs_numeric)
    iugs_numeric_filtered <- iugs_numeric[valid_iugs_idx]
    df_iugs_filtered <- df[valid_iugs_idx, ]
    
    upper_value_candidates <- iugs_numeric_filtered[iugs_numeric_filtered > interp_value]
    
    if (length(upper_value_candidates) == 0) {
      cat(sprintf('IUGS 2023.9列中没有大于插值值%.4f的数据\n', interp_value))
      
      if (length(iugs_numeric_filtered) > 0) {
        # 找最接近的值
        closest_idx <- which.min(abs(iugs_numeric_filtered - interp_value))
        upper_row_for_table <- df_iugs_filtered[closest_idx, ]
        upper_value_display <- upper_row_for_table[[iugs_col]]
        cat(sprintf("尝试使用IUGS 2023.9中最接近的值: %s\n", upper_value_display))
      } else {
        return(NULL)
      }
    } else {
      upper_value_final <- min(upper_value_candidates)
      upper_idx_final <- which(iugs_numeric_filtered == upper_value_final)[1]
      
      if (is.na(upper_idx_final)) {
        cat(sprintf("无法找到值为 %s 的原始索引。\n", upper_value_final))
        return(NULL)
      }
      
      upper_row_for_table <- df_iugs_filtered[upper_idx_final, ]
      upper_value_display <- upper_value_final
    }
    
    if (is.null(upper_row_for_table)) {
      cat("无法确定用于插值显示的对应行。\n")
      return(NULL)
    }
    
    upper_idx_original <- which(valid_iugs_idx)[which(rownames(df_iugs_filtered) == rownames(upper_row_for_table))[1]]
    result_df <- df[upper_idx_original, show_cols, drop = FALSE]
    result_df[[version]] <- age_num
    result_df[[iugs_col]] <- interp_value
    
    labels <- c("宇", "界", "系", "统", "阶")
    info_row <- df[upper_idx_original, ]
    info <- character(0)
    
    for (i in 1:length(labels)) {
      col_label <- labels[i]
      if (col_label %in% names(df)) {
        info <- c(info, paste0(info_row[[col_label]], labels[i]))
      }
    }
    
    cat(sprintf("插值结果: %s IUGS插值:%.4f (原上界IUGS:%s, 输入年代:%s)\n", 
                paste(info, collapse = " "), interp_value, upper_value_display, age_num))
    
    return(result_df)
  }
}

# 主函数
main <- function() {
  # 读取Excel文件
  df <- tryCatch({
    read_excel('C:/Users/Administrator/Desktop/WJY2/1937.xlsx', col_types = "text")  # 确保所有列作为文本读取
  }, error = function(e) {
    cat(sprintf("读取wjy.xlsx失败: %s\n", e$message))
    cat("尝试查找当前目录下的Excel文件...\n")
    excel_files <- list.files(pattern = "\\.xlsx$")
    if (length(excel_files) > 0) {
      cat(sprintf("找到以下Excel文件:\n%s\n", paste(excel_files, collapse = "\n")))
      cat("请确保wjy.xlsx存在于当前目录下\n")
    } else {
      cat("当前目录下没有找到Excel文件\n")
    }
    return(NULL)
  })
  
  if (is.null(df)) {
    return(NULL)
  }
  
  # 显示数据表的基本信息
  cat(sprintf("\nwjy.xlsx基本信息:\n行数: %d, 列数: %d\n", nrow(df), ncol(df)))
  
  # 过滤掉空白列、带有"误差"的列和以'Unnamed:'开头的列
  valid_cols <- names(df)[sapply(names(df), function(col) {
    return(col != "" && !grepl("误差", col) && !grepl("^Unnamed:", col))
  })]
  df <- df[, valid_cols]
  
  # 显示所有可选年代表版本
  cat('可选年代表版本（来自主数据表 wjy.xlsx）：\n')
  cat(paste(names(df), collapse = ' | '), '\n\n')
  
  version <- readline(prompt = '输入当前年代表版本（主数据表中的列名）：')
  version <- trimws(version)
  
  if (!(version %in% names(df))) {
    cat(sprintf('未找到年代表版本：%s\n', version))
    return(NULL)
  }
  
  # 使用命令行输入文件路径
  cat("\n请输入包含年代数据的Excel文件路径：\n")
  selected_excel_file_path <- readline(prompt = "文件路径: ")
  selected_excel_file_path <- trimws(selected_excel_file_path)
  
  if (selected_excel_file_path == "") {
    cat("未输入任何文件路径。\n")
    return(NULL)
  }
  
  cat(sprintf("\n您输入的文件路径是: %s\n", selected_excel_file_path))
  
  # 检查文件是否存在
  if (!file.exists(selected_excel_file_path)) {
    cat(sprintf("错误：文件 %s 不存在。\n", selected_excel_file_path))
    return(NULL)
  }
  
  ages_from_file <- c()
  age_column_name <- ""
  source_excel_df_for_read <- NULL
  
  tryCatch({
    source_excel_df_for_read <- read_excel(selected_excel_file_path, col_types = "text")  # 确保所有列作为文本读取
    file_basename <- basename(selected_excel_file_path)
    cat(sprintf("\n已读取文件: %s\n", file_basename))
    cat("文件中的列名：\n")
    cat(paste(names(source_excel_df_for_read), collapse = ' | '), '\n\n')
    
    age_column_name <- readline(prompt = "请输入包含年代数据的那一列的名称：")
    age_column_name <- trimws(age_column_name)
    
    if (!(age_column_name %in% names(source_excel_df_for_read))) {
      cat(sprintf("在 %s 中未找到列：%s\n", file_basename, age_column_name))
      return(NULL)
    }
    
    raw_ages <- as.character(source_excel_df_for_read[[age_column_name]])
    ages_from_file <- trimws(raw_ages[trimws(raw_ages) != ""])
    
  }, error = function(e) {
    cat(sprintf("读取或处理Excel文件时发生错误: %s\n", e$message))
    return(NULL)
  })
  
  if (length(ages_from_file) == 0) {
    cat(sprintf("在文件 %s 的 '%s' 列中未找到有效年代数据。\n", basename(selected_excel_file_path), age_column_name))
    return(NULL)
  }
  
  iugs_col <- 'IUGS 2023.9'
  show_cols <- c('宇', '界', '系', '统', '阶', version, iugs_col)
  show_cols <- show_cols[show_cols %in% names(df)]
  
  results_dfs <- list()
  processed_iugs_map <- list()
  
  for (age_str in ages_from_file) {
    result_df <- process_age_data(age_str, df, version, iugs_col, show_cols)
    if (!is.null(result_df)) {
      results_dfs[[length(results_dfs) + 1]] <- result_df
      if (iugs_col %in% names(result_df) && nrow(result_df) > 0) {
        processed_iugs_map[[age_str]] <- result_df[[iugs_col]][1]
      }
    }
  }
  
  if (length(results_dfs) == 0) {
    cat("\n所有输入年代均未能成功处理，无法生成。\n")
  } else {
    final_table_df <- bind_rows(results_dfs)
    
    # 写回处理结果到原始Excel文件
    if (!is.null(source_excel_df_for_read) && age_column_name != "") {
      tryCatch({
        df_to_write_back <- source_excel_df_for_read
        
        # 创建新的IUGS数据列
        new_iugs_column_data <- rep(NA, nrow(df_to_write_back))
        original_age_values <- as.character(df_to_write_back[[age_column_name]])
        
        for (i in 1:length(original_age_values)) {
          age_val <- trimws(original_age_values[i])
          if (age_val != "" && !is.null(processed_iugs_map[[age_val]])) {
            new_iugs_column_data[i] <- processed_iugs_map[[age_val]]
          }
        }
        
        # 确定插入位置和新列名
        insert_loc <- which(names(df_to_write_back) == age_column_name) + 1
        new_column_name <- paste0(iugs_col, " (处理结果)")
        
        # 如果新列名已存在，则添加后缀
        temp_new_col_name <- new_column_name
        counter <- 1
        while (temp_new_col_name %in% names(df_to_write_back)) {
          temp_new_col_name <- paste0(new_column_name, "_", counter)
          counter <- counter + 1
        }
        new_column_name <- temp_new_col_name
        
        # 插入新列
        df_to_write_back <- df_to_write_back %>%
          add_column(!!new_column_name := new_iugs_column_data, .after = age_column_name)
        
        # 写回文件
        write_xlsx(df_to_write_back, selected_excel_file_path)
        cat(sprintf("\n处理结果已成功写入到文件 '%s' 的 '%s' 列。\n", 
                   basename(selected_excel_file_path), new_column_name))
        
      }, error = function(e) {
        cat(sprintf("将结果写回到Excel文件时发生错误: %s\n", e$message))
      })
    }
  }
}

# 运行主函数
main() 