def create_line_chart(df):
    """Create a line chart with better separation between lines"""
    plt.figure(figsize=(14, 8))
    
    # 修正后的列表推导式
    plot_columns = [
        col for col in df.columns 
        if (pd.api.types.is_numeric_dtype(df[col]) and 
           not all(df[col].fillna(0) == 0))
    ]
    
    if not plot_columns:
        return None
    
    # 确保索引是datetime类型
    try:
        df.index = pd.to_datetime(df.index)
    except Exception as e:
        print(f"日期转换错误: {e}")
        return None
    
    # 创建图表 - 使用不同的线型和标记增强区分度
    line_styles = ['-', '--', '-.', ':']
    markers = ['o', 's', '^', 'D', 'v', 'p', '*', 'h']
    colors = plt.cm.tab10.colors  # 使用标准颜色循环
    
    ax = plt.gca()
    
    # 计算合适的y轴范围
    min_rate = df[plot_columns].min().min()
    max_rate = df[plot_columns].max().max()
    rate_range = max_rate - min_rate
    
    # 如果数据范围太小，放大显示区域
    y_padding = max(0.1, rate_range * 0.5)  # 至少0.1%的padding
    plt.ylim(min_rate - y_padding, max_rate + y_padding)
    
    # 绘制每条线
    for i, column in enumerate(plot_columns):
        # 循环使用不同的线型、标记和颜色
        line_style = line_styles[i % len(line_styles)]
        marker = markers[i % len(markers)]
        color = colors[i % len(colors)]
        
        line = ax.plot(
            df.index, 
            df[column], 
            label=column,
            linestyle=line_style,
            marker=marker,
            color=color,
            linewidth=2,
            markersize=8,
            markeredgecolor='white',
            markeredgewidth=1
        )
        
        # 添加数据标签（可选）
        if len(df) <= 7:  # 数据点不多时才显示标签
            for x, y in zip(df.index, df[column]):
                ax.annotate(
                    f"{y:.4f}%",
                    (x, y),
                    textcoords="offset points",
                    xytext=(0, 10),
                    ha='center',
                    fontsize=9,
                    bbox=dict(
                        boxstyle='round,pad=0.2',
                        fc='white',
                        alpha=0.7
                    )
                )
    
    # 图表美化
    ax.set_title('Japanese Yen TIBOR Rates (Enhanced Visibility)', fontsize=14, pad=20)
    ax.set_ylabel('Rate (%)', fontsize=12)
    ax.set_xlabel('Date', fontsize=12)
    
    # 日期格式设置
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
    ax.xaxis.set_major_locator(mdates.AutoDateLocator())
    plt.xticks(rotation=45, ha='right')
    
    # 网格和背景
    ax.grid(True, linestyle='--', alpha=0.6)
    ax.set_axisbelow(True)
    
    # 图例位置调整
    ax.legend(
        bbox_to_anchor=(1.05, 1),
        loc='upper left',
        frameon=True,
        framealpha=1,
        edgecolor='#333',
        title='Tenor',
        title_fontsize=12
    )
    
    # 调整布局
    plt.tight_layout()
    
    # 保存图表
    buf = io.BytesIO()
    plt.savefig(
        buf, 
        format='png', 
        dpi=150, 
        bbox_inches='tight',
        facecolor='white'
    )
    buf.seek(0)
    plt.close()
    
    return buf
