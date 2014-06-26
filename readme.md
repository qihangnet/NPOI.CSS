NPOI.CSS用法
------------------
目前，本扩展只支持.NET4及以上版本的项目，低版本的暂时不支持，请见谅。

1. 引用NPOI.CSS.dll
2. `using NPOI.CSS;`
3. `cell.CSS("color:red;font-weight:bold;font-size:11;font-name:宋体;border-type:thin;")`

## 设置字体 font-name   
	注意只能有一个值

## 字体下划线 font-underline  
-	none	无下划线
-	single	单下划线
-	double	双下划线
-	single_accounting	会计用单下划线
-	double_accounting	会计用双下划线
## 字标 font-superscript
-	none	无
-	sub	下标
-	super	上标	

## 横向对齐方式 text-align
-	general
-	left
-	center
-	right
-	fill
-	justify
-	center_selection
-	distributed

## 纵向对齐方式 vertical-align
-	top
-	center
-	bottom
-	justify
-	distributed


## 边框样式 border-type
-	none
-	thin
-	medium
-	dashed
-	hair
-	thick
-	double
-	dotted
-	medium_dashed
-	dash_dot
-	medium_dash_dot
-	dash_dot_dot
-	medium_dash_dot_dot
-	slanted_dash_dot

## 斜体 font-italic
-	true
-	false

## 删除线 font-strikeout
-	true
-	false

## 数据格式 data-format

	参考NPOI官方文档中格式化字符串的设定要求

## 字体颜色 color / 背景色 background-color
-	black
-	white
-	red
-	bright_green
-	blue
-	yellow
-	pink
-	turquoise
-	dark_red
-	green
-	dark_blue
-	dark_yellow
-	violet
-	teal
-	grey_25_percent
-	grey_50_percent
-	cornflower_blue
-	maroon
-	lemon_chiffon
-	orchid
-	coral
-	royal_blue
-	light_cornflower_blue
-	sky_blue
-	light_turquoise
-	light_green
-	light_yellow
-	pale_blue
-	rose
-	lavender
-	tan
-	light_blue
-	aqua
-	lime
-	gold
-	light_orange
-	orange
-	blue_grey
-	grey_40_percent
-	dark_teal
-	sea_green
-	dark_green
-	olive_green
-	brown
-	plum
-	indigo
-	grey_80_percent
-	automatic

为了提升输入效率，cssKey支持缩写别名：
---------------------------------------------	  
-	**font-color**  
	color  
	fc  
-	**font-weight**  
	fw  
-	**font-name**  
	fn  
-	**font-size**  
	fs
-	**font-italic**  
	italic  
	i    
-	**font-underline**  
	underline    
	u  	
-	**font-strikeout**  
	deleteline  
	d-line  
	strikeout  
	d  
-	**font-superscript**  
	font-offset  
	superscript  
	fss  
	ss  
-	**background-color**  
	bg-color  
	bg-c  
	bgc  
-	**text-align**  
	align  
-	**vertical-align**  
	v-align
-	**data-format**  
	format
