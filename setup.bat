@if(0)==(0) ECHO OFF
cscript.exe //nologo //E:JScript "%~f0" %*
GOTO :EOF
@end

var globals = (function() {
  return {
    shell : WScript.CreateObject("WScript.Shell"),
    fsys  : WScript.CreateObject("Scripting.FileSystemObject")
  };
})();

var ssfLOCALAPPDATA = 28;
var localApp = WScript.CreateObject("Shell.Application").Namespace(ssfLOCALAPPDATA).Self.Path;
var desktop = globals.shell.SpecialFolders("Desktop");
var source = globals.fsys.BuildPath(globals.fsys.getParentFolderName(WScript.ScriptFullName), 'csvhub.ps1')
var dest = globals.fsys.BuildPath(localApp, Math.floor( new Date().getTime() / 1000 ) + '.ps1');

globals.fsys.CopyFile(source, dest);

var shortcut = globals.shell.CreateShortcut(globals.fsys.BuildPath(desktop, 'ここへドロップ.lnk'));
shortcut.TargetPath = "powershell.exe";
shortcut.Arguments = "-ExecutionPolicy RemoteSigned -NoProfile -File " + dest;
shortcut.WindowStyle = 7;// 最小化
shortcut.Save();
shortcut = null;
WScript.Echo('インストール完了');
