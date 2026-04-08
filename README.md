# Comparison between Drones and Trucks in Cross-Provincial Delivery

本项目用于比较无人机与货车在跨区域运输场景下的运输效率差异，并通过批量模拟输出不同重量、不同距离区间下两种运输方式的时间对比结果。

## 项目功能

本项目主要实现以下功能：

- 随机生成订单场景
- 基于城市、蜂巢节点和随机起终点模拟运输任务
- 计算无人机运输时间
- 基于离线 OSM 路网计算货车运输时间
- 比较无人机与货车在不同重量下的效率差异
- 输出批量模拟结果和汇总结果

## 主要文件

- `无人机货车对比.py`：主程序
- `requirements.txt`：Python 依赖包说明
- `README.md`：项目说明文档

## 运行环境

建议使用 Python 3.10 及以上版本。

本项目依赖以下 Python 库：

- pandas
- osmnx
- networkx
- openpyxl

## 输入数据要求

程序运行前需要准备以下数据文件：

### 1）Excel 数据文件
代码中使用的 Excel 文件名为：


最新的快递量分析_补充蜂巢城市经纬度.xlsx
其中 Sheet1 至少需要包含以下列：

城市
纬度
经度
蜂巢城市
蜂巢城市中心纬度
蜂巢城市中心经度

### 2）离线 OSM 数据文件

代码中使用的离线地图文件为：

china-260401.osm.pbf

该文件用于配合 osmium 进行路网裁切和货车路线计算。

## 运行前需要修改的路径

请根据你自己的电脑环境，修改代码中的以下路径：

BASE_DIR
EXCEL_PATH
OSM_PBF_PATH
OSMIUM_CMD

例如：

BASE_DIR = Path(r"C:\Users\你的用户名\项目路径")
## 安装依赖

在项目目录下运行：

pip install -r requirements.txt
## 运行方式

在终端中进入项目目录后执行：

python 无人机货车对比.py
## 输出结果

程序运行后会输出：

每单模拟结果
无人机与货车时间对比结果
按重量汇总的统计结果
Excel 输出文件
## 项目说明

当前模型的重点在于从仿真角度比较无人机与货车在典型跨区域运输任务中的效率差异。
结果通常不能简单理解为“无人机全面替代货车”，而更适合用于识别无人机在哪些重量、距离和场景下具备更高效率。

## 注意事项
本项目使用离线 OSM 路网数据，因此需要提前准备 .osm.pbf 文件。
代码中存在本地绝对路径，上传到 GitHub 后，其他人运行前需要先修改路径。
若仓库中未上传大体积数据文件，使用者需要自行准备 Excel 和 OSM 数据文件。
由于使用 .osm.pbf 文件较大，因此运行较慢。
## 作者

GitHub: https://github.com/shiyueqinga/Comparison-between-Drones-and-Trucks-in-Cross-Provincial-Delivery
