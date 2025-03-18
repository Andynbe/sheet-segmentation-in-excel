# Excel 工作表分割汇总工具

#### 项目介绍
基于xlwings开发的Excel自动化处理工具，提供以下核心功能：
- 批量处理指定目录下的多个Excel文件
- 智能识别工作表结构
- 按指定工作表序号合并数据
- 自动保留源文件标识
- 生成带时间戳的汇总文件

#### 软件架构
- 核心库：xlwings 0.30.12+
- 开发语言：Python 3.8+
- 依赖管理：pip
- 目录结构：
  ```
  ├── data/         # 原始数据存放目录
  ├── output/       # 结果输出目录
  ├── main_v1.5.py  # 主程序
  └── utils.py      # 路径处理工具
  ```

#### 安装教程
1. 环境要求
   - Windows 10/11 系统
   - Microsoft Office 2016+
   - Python 3.8+

2. 依赖安装
   ```powershell
   pip install xlwings==0.30.12
   pip install pywin32==306
   ```

3. 项目配置
   - 在项目根目录创建data文件夹存放待处理Excel
   - 确保所有Excel文件为xlsx格式

#### 使用说明
1. 文件准备
   - 将需要处理的Excel文件放入data目录
   - 保持Excel文件结构一致（相同列结构）

2. 运行程序
   ```powershell
   python main_v1.5.py
   ```

3. 操作流程
   - 程序将显示检测到的工作表列表
   - 输入需要处理的sheet序号（数字）
   - 等待处理完成提示（约1文件/秒）
   - 结果文件自动生成在项目根目录

4. 输出示例
   ```
   应付账款_汇总结果_20240318.xlsx
   └── Sheet1
       ├── A1:H100  源数据
       └── I列       源文件名标识
   ```

#### 注意事项
1. 文件规范
   - 单个文件建议不超过50万行
   - 合并后总行数不超过Excel限制（1048576行）
   - 文件名建议使用英文命名

2. 异常处理
   - 遇到程序中断时，请检查：
     - Excel文件是否被其他程序占用
     - 工作表序号是否输入正确
     - data目录是否存在且不为空

3. 性能优化
   - 处理万行级文件时，建议关闭其他Excel实例
   - 如需处理特大文件，可联系开发者获取专业版

#### 版本更新
v1.5 更新内容：
- 增加进度提示功能
- 优化文件路径处理
- 修复多sheet识别异常
- 增强错误输入校验
