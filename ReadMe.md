## 从excel中导出联系人发送QQ邮件
- 需提前在发送方邮箱中进行设置，讲smtp服务器发送功能开启
- python 3.x
- 需加载模块(在安装有pip的前提下)

```
pip install xlrd

```

### 需注意替换的参数
- 服务器（此处为QQsmtp 服务器 使用默认端口25）
- 用户名
- 授权码（通过邮箱自动生成的随机码）
- 发送内容
- 发送方昵称
- 接收方昵称
- 根据表格具体情况设置需获取单元格位置

### 调试建议
- 注意保持‘utf-8’编码
- 发送方和接收方可能需要经过`Header`处理
- 表格内容需要通过`value`属性获取