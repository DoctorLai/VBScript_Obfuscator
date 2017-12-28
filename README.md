# VBScript Obfuscator/Defuscator in VBScript
The VBScript Obfuscator written in [VBScript](https://isvbscriptdead.com/vbs-obfuscator/)

# How to Use?
The command line utility is interpreted under Window Scripting Host. VBScript source files can be passed as command line parameters and the obfuscated source will be printed to console.

`cscript.exe vbs_obfuscator.vbs sample.vbs`

![](https://github.com/DoctorLai/VBScript_Obfuscator/blob/master/vbscript-obfuscated-in-vbscript.png?raw=true)


# Defuscator in VBScript
You can revert the process via `vbs_defuscator.vbs`

```
cscript.exe vbs_defuscator.vbs sample_defuscated.vbs > sample.vbs
```

# ROT47 Obfuscator
This project provides another VBScript obfuscator based on [ROT47](https://rot47.net), see the code [vbs_rot47_obfuscator.vbs](https://github.com/DoctorLai/VBScript_Obfuscator/blob/master/vbs_rot47_obfuscator.vbs)

# BASE64 Obfuscator
This project provides another VBScript obfuscator based on [BASE64](https://rot47.net/base64encoder.html), see the code [vbs_obfuscator_base64.vbs](https://github.com/DoctorLai/VBScript_Obfuscator/blob/master/vbs_obfuscator_base64.vbs)

# Author
[@justyy](https://steemit.com/@justyy)
