# Yuanyue_SP1PI
<img width="1765" height="571" alt="sp1piLogo" src="https://github.com/user-attachments/assets/78a60ab1-5d8e-4e91-b436-fb0a599439fd" />

Yuanyue Sp1pi/源悦-石皮皮 

一个自动补丁安装器，专用于[源悦TTS。](https://github.com/CN-Air84/YuanYue-TTS/)

大部分代码支持几乎所有Win7及以后的系统，但installer.vbs里写了限制，只能在win7 x64系统中使用。

你可以任意修改程序，使程序支持任意32位系统/Win8/Win8.1/Win10/……。

许可证是我瞎选的，不用管。

不允许以任何形式用于任何形式的商业软件。

非商业软件的话，随便你怎么搞都可以，代码最后端留下本仓库的链接即可。 

短时间内石皮皮不会再有更新了。

exe编译使用bat2exe、vbs2exe、单文件制作程序等。

把需要安装的补丁命名为7patch_*******.msu，然后移到和install.vbs相同的路径下。

#### 怎么是全英文的啊？

一切服务于兼容性。

要不是为了这个我早拿python写了。

vbscript反正我是再也不会碰了。运行都不运行的那种。我自己写了不下四版，又让deepseek给我修bug，修了七版，他没修好。又交gemini改了三版，没修好。又交ds，可算能跑了。太tm折磨了。

#### 如何打包？

第一步，想办法把vbs转成exe。

第二步（可选），如果你的程序只能给32位系统用，想办法把32-only打包成exe，

反之，只能给64位系统用，就把64-only打成exe。

然后塞到和install.vbs同目录下。

第三步，把需要安装的补丁命名为7patch_*******.msu，然后移到和install.vbs相同的路径下。

第四步。[下载单文件打包程序。我用的是这个，别的什么五花八门的也都可以。我就拿这个示例了。](https://www.52pojie.cn/thread-1696725-1-1.html)

然后把你存有两个exe和一大堆补丁的路径输进去，设置好x86、x64运行。

完事。
