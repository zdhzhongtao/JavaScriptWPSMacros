<!--
 * @Author: Wade Zhong wzhong@hso.com
 * @Date: 2024-08-12 09:09:21
 * @LastEditTime: 2024-08-12 09:10:32
 * @LastEditors: Wade Zhong wzhong@hso.com
 * @Description: 
 * @FilePath: \PythonSolution\README.md
 * Copyright (c) 2024 by Wade Zhong wzhong@hso.com, All Rights Reserved. 
-->
# Excel 笛卡尔积生成器

这个 Python 脚本用于读取 Excel 文件中的多个 sheet，计算它们的笛卡尔积，并将结果保存到同一文件的新 sheet 中。

## 功能

- 读取指定 Excel 文件的所有 sheet（除了名为 "Result" 的 sheet）
- 计算所有 sheet 数据的笛卡尔积
- 显示处理进度和预计剩余时间
- 将结果保存到名为 "Result" 的新 sheet 中
- 输出总处理时间

## 安装

1. 确保你的系统已安装 Python 3.6 或更高版本。

2. 克隆或下载此仓库到你的本地机器。

3. 在项目目录中打开命令行，运行以下命令安装所需的依赖：

   ```
   pip install -r requirements.txt
   ```

## 使用方法

1. 打开命令行，导航到脚本所在的目录。

2. 运行以下命令：

   ```
   python excel_cartesian_product.py
   ```

3. 根据提示输入 Excel 文件的完整路径。

4. 脚本将处理文件并在同一文件中创建或更新 "Result" sheet。

## 注意事项

- 确保 Excel 文件未被其他程序打开。
- 请确保你有足够的权限访问和修改指定的 Excel 文件。
- 对于大型数据集，处理可能需要较长时间。请耐心等待，进度条会显示估计的剩余时间。

如果你遇到任何问题或有改进建议，请提出 issue 或 pull request。