# jdrawing
Solidworks PMR Framework
==== 
![image](https://github.com/jinhefeng/jdrawing/blob/master/images/PMR.png)
本项目致力于创造一个程序、模型、参数化规则进行分离的二次开发世界，使程序员专注于程序、制图员专注于模型、机械工程师专注于参数化规则，从而各取所长良性发展。

## 应用举例
> ##### 第一步，建立三维模版，在方程式中建立变量和尺寸的对应关系
![image](https://github.com/jinhefeng/jdrawing/blob/master/images/test_size_modify_sldprt.png)
> ##### 第二步，打开文件，使用Setsize(方程式变量名，尺寸数值)修改尺寸
![image](https://github.com/jinhefeng/jdrawing/blob/master/images/test_size_modify.png)
> ##### 第三步，commit()

## 发展计划
> #### 1. 完善JSolidworks操作类
>> ##### 1.1 对零部件尺寸的修改
>> ##### 1.2 对零部件特征的新建、压缩及还原
>> ##### 1.3 对部件进行装配
>> ##### 1.4 对工程图基于视图微调
> #### 2. 完善Solidworks解释器
>> ##### 2.1 建立解释器脚本语法
>> ##### 2.3 脚本自动生成器及excel接口
> #### 3. 完善Solidworks建模工具
>> ###### 3.1 快速定义方程式
>> ###### 3.2 快速导出解释器脚本

## 注意事项：
1. 本项目致力于让普通机械工程师也能进行Solidwork二次开发，有时为了操作的简便或许会牺牲一定的效率；
2. 本项目首次提出了PMR框架，并在接下来的工作中请各位Contrubitor始终坚持这一原则；
3. 本项目暂定基于Solidworks 2016开发
