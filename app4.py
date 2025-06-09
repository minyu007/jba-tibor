def split_row_to_rows(df):
    if df.empty:
        return pd.DataFrame()
    
    first_row = df.iloc[0].copy()
    first_row = first_row.replace({np.nan: ''})
    split_data = {}
    
    # 首先处理分割数据
    for col in first_row.index:
        if first_row[col] == '':
            split_data[col] = ['']
        else:
            # 分割字符串并去除空白字符
            split_values = [v.strip() for v in first_row[col].split('\r')]
            split_data[col] = split_values
    
    max_length = max(len(v) for v in split_data.values())
    
    # 统一长度并转换数字
    for col in split_data:
        current_length = len(split_data[col])
        if current_length < max_length:
            split_data[col].extend([''] * (max_length - current_length))
        
        # 尝试将非空字符串转换为float
        converted_values = []
        for val in split_data[col]:
            if val == '':
                converted_values.append(np.nan)
            else:
                try:
                    # 移除可能存在的百分号或其他非数字字符
                    cleaned_val = val.replace('%', '').strip()
                    converted_values.append(float(cleaned_val))
                except (ValueError, AttributeError):
                    # 如果转换失败，保留原始值（通常是日期列）
                    converted_values.append(val)
        split_data[col] = converted_values
    
    new_df = pd.DataFrame(split_data)
    
    # 确保所有数值列都是float类型
    for col in new_df.columns:
        if col != 'Date' and new_df[col].dtype == 'object':
            try:
                new_df[col] = pd.to_numeric(new_df[col], errors='ignore')
            except:
                pass
    
    return new_df
