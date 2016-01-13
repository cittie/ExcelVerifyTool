运行环境：
python 2.7以上版本，包括3.x
OSX 或 Windows环境
需要xlrd库

config.ini配置说明：

[DEFAULT]    基础配置，名称不可修改
COLUMN_TITLE_LINE = 1    表标题所在行数
COLUMN_CONTENT_LINE = 3      表内容开始行数
PROJECT_PATH = ../../DebugProj/      工程所在目录
EXCEL_PATH = ../../Excel/        表所在目录

[Item ID]        检查项名字，不可重名
check_type = id      检查类型，可以为: id, group, file，ids 不同的类型后续字段有区别
source = PQ_Item         来源文件名，string，仅一个
source_sheet =      可选 来源表名，string，仅一个
source_title = ID WalletID ToolID        来源数据标题， string，一个或多个，用空格分隔
target = PQ_Loot.xlsx PQ_Gear.xlsx         目标文件名，string，一个或多个，用空格分隔
target_title = Item001 Item002 Item003 EvolveMaterial EvolveMaterialReal        目标数据标题， string，一个或多个，用空格分隔

****以下是检查类型为file的项****
pre_path = Assets/Resources/PQCharacters/Pet_Cat/        目标文件在工程中的所在目录
extension_name =            选填 目标文件扩展名，不区分大小写

****以下是检查类型为ids的项****
eigen_string = IDS_          内容开头的特征字符串

****以下是检查类型为group的项****
group_sheet = SDLootItem.xlsx            文件名，string，仅一个
group_title = ID             来源数据标题， string，一个或多个，用空格分隔
