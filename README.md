vbs-guid
========

VBS入门到精通


#scripting.filesystemobject

-------------

*作者：zhangbo2012



*QQ： 369029696



*blog：http://www.cnblogs.com/zhangbo2012



本教程为原创内容，转载请注明出处



--------------



##简介

 scripting.filesystemobject提供访问文件系统的服务，它是一个系列对象的集合。主要用于实现文件读写，文件目录操作，磁盘管理等功能。在使用时一般是先创建FSO对象，然后获取子对象（如文件，文件夹，驱动器），再对这些子对象进行操作。

 如下图所示：

 

 其常用的对象和集合有


 <table>

<tr><td>名称</td><td>类型</td><td>说明</td></tr>

<tr><td>Drive</td><td>对象</td><td>驱动器对象，提供硬盘分区、CD-ROM、和U盘等驱动器信息</td></tr>

<tr><td>Drives</td><td>集合</td><td>驱动器列表</td></tr>

<tr><td>File</td><td>对象</td><td>文件对象，用于查看文件信息和执行文件操作。</td></tr>

<tr><td>Files</td><td>集合</td><td>提供包含在文件夹内的所有文件的列表。</td></tr>

<tr><td>Folder</td><td>对象</td><td>文件夹对象，用于查看文件夹信息和执行文件夹操作。</td></tr>

<tr><td>Folders</td><td>集合</td><td>提供在 Folder 内的所有文件夹的列表。</td></tr>

<tr><td>TextStream</td><td>对象</td><td>用来读写文本文件。</td></tr>

</table>



##Drives列表和Drive对象

###示例

  下面这个脚本可以显示当前操作系统上所有可用的驱动器，及其属性：

 ```

 'code by zhangbo2012

 '-----------------------------------------------

 dim fso                                           

 set fso=createobject("scripting.filesystemobject")

 set Drives=fso.Drives  

 ol=""                           

 for each drive in Drives                          

 	if drive.IsReady then 

		ol = ol & vbcrlf & " 路径           " & drive.Path

		ol = ol & vbcrlf & " 根目录         " & drive.RootFolder                           

		ol = ol & vbcrlf & " 卷名称         " & drive.VolumeName

		ol = ol & vbcrlf & " 总空间         " & drive.TotalSize

		ol = ol & vbcrlf & " 可用空间       " & drive.FreeSpace

		ol = ol & vbcrlf & " 类型           " & drive.DriveType

		ol = ol & vbcrlf & " 文件系统       " & drive.FileSystem

		ol = ol & vbcrlf & " SerialNumber   " & drive.SerialNumber

		ol = ol & vbcrlf & string(50,"-")  

	end if 

 next     

 wscript.echo ol 


 ```

 * 第3-4行 创建fso对象；

 * 第5行 通过fso对象创建drivers列表；

 > drivers列表可以认为是由多个drive对象组成的数组

 * 第7-19行 收集每一个drvie对象的常用属性

 >drive.IsReady表示当前driver是否可用，不可用的driver会有很多属性访问不了



在我的电脑上运行结果如下

```

E:\Program\VBS\教程\script\04-3 filesystemobject>cscript.exe driver.vbs

Microsoft (R) Windows Script Host Version 5.8

版权所有(C) Microsoft Corporation 1996-2001。保留所有权利。





 路径           C:

 根目录         C:\

 卷名称

 总空间         42952376320

 可用空间       8962809856

 类型           2

 文件系统       NTFS

 SerialNumber   1379193556


--------------------------------------------------

 路径           D:

 根目录         D:\

 卷名称

 总空间         85913014272

 可用空间       28353851392

 类型           2

 文件系统       NTFS

 SerialNumber   726565

--------------------------------------------------

 路径           E:

 根目录         E:\

 卷名称

 总空间         85913014272

 可用空间       18478673920

 类型           2

 文件系统       NTFS

 SerialNumber   345843

--------------------------------------------------

 路径           F:

 根目录         F:\

 卷名称         新加卷

 总空间         51539603456

 可用空间       38677069824

 类型           2

 文件系统       NTFS

 SerialNumber   1424353221

--------------------------------------------------


```



> 从上面的运行结果可以看出，在展示空间时用的单位字节，可读性不高。

为了提升可读性，我们可以将其按  1024字节=1MB 1024MB=1GB的换算规则算成以GB为单位进行展示。

请各位同学可以尝试对脚本进行优化。





##folder对象

###简介

  folder 即文件夹对象，可以实现读取文件夹属性和对文件夹进行移动、复制、删除等操作。

###folder对象属性

<table>

<tr><td>属性</td><td>说明</td></tr>

<tr><td>Attributes</td><td><table>标识文件或者文件夹属性
<tr><td>常数</td><td>值</td><td>描述</td></tr>

<tr><td>Normal</td><td>0</td><td>普通文件。不设置属性。</td></tr>

<tr><td>ReadOnly</td><td>1</td><td>只读文件。属性为读/写。</td></tr>

<tr><td>Hidden</td><td>2</td><td>隐藏文件。属性为读/写。</td></tr>

<tr><td>System</td><td>4</td><td>系统文件。属性为读/写。</td></tr>

<tr><td>Volume</td><td>8</td><td>磁盘驱动器卷标。属性为只读。</td></tr>

<tr><td>Directory</td><td>16</td><td>文件夹。属性为只读。</td></tr>

<tr><td>Archive</td><td>32</td><td>文件在上次备份后已经修改。属性为读/写。</td></tr>

<tr><td>Alias</td><td>64</td><td>链接或者快捷方式。属性为只读。</td></tr>

<tr><td>Compressed</td><td>128</td><td>压缩文件。属性为只读。</td></tr>

</table></td></tr>

<tr><td>DateCreated</td><td>创建日期</td></tr>

<tr><td>DateLastAccessed</td><td>最后访问日期</td></tr>

<tr><td>DateLastModified</td><td>最后修改日期</td></tr>

<tr><td>Drive</td><td>驱动器</td></tr>

<tr><td>IsRootFolder</td><td>是否为根目录</td></tr>

<tr><td>Name</td><td>名称</td></tr>

<tr><td>ParentFolder</td><td>父目录</td></tr>

<tr><td>Path</td><td>路径</td></tr>

<tr><td>ShortName</td><td>ShortName</td></tr>

<tr><td>ShortPath</td><td>ShortPath</td></tr>

<tr><td>Size</td><td>占用空间</td></tr>

<tr><td>Type</td><td>类型</td></tr>

<tr><td>SubFolders</td><td>子文件夹集合</td></tr>

<tr><td>Files</td><td>子文件集合</td></tr>

</table>

###folder对象方法

<table>

<tr><td>方法</td><td>说明</td></tr>

<tr><td>Copy</td><td>复制</td></tr>

<tr><td>Delete</td><td>删除</td></tr>

<tr><td>Move</td><td>移动</td></tr>

</table>







###示例

  下面 其属性：




###示例

####读文件

1. 使用方法
``object.Run(strCommand, [intWindowStyle], [bWaitOnReturn]) ``
 * strCommand 命令行，必填字段，为程序路径，也可以在后面加上调用的参数
 * intWindowStyle 打开程序的窗口状态，可选字段，默认为显示并激活窗口，取值有：


##示例

###示例1

  xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

 ```  xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx```



###示例2
