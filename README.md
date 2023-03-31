# VirtualAudioRegister
Registers a virtual audio device that captures what you hear from the speakers.

The device is a DirectShow filter and can be used with FFmpeg to record the audio.
Developed by Roger Pack.

https://github.com/rdp/virtual-audio-capture-grabber-device

VirtualAudioRegister.exe can be used to register or un-register the filters.
This has Administrator privileges and Windows UAC might intercept, but just allow it.

If "virtual-audio-capturer" has not been registered, click the 'Register' button. 
This will register both 32 bit and 64 bit versions.
After confirmation of success, the button will show 'UnRegister'.
Click 'Unregister' to remove it from the system. To update to new dll files, click 'Unregister' and then 'Register' again.

### Manual registration

You can also register manually. There are separate folders for 32 bit and 64 bit versions and you can install both on a 64 bit system. In each folder you will find "_register_run_as_admin.bat". RH click and "Run as Administrator" to register. Register the 32 bit version first.

### Building the project

The project files are for Visual Studio 2022
The application should be built x64 to register both versions directly using the registry.
A 64 bit Windows system is required. _WIN64 is checked at compile, not at runtime
If built with Visual Studio "C++ > Code Generation > Runtime Library /MT",
the Visual Studio runtime dlls are not necessary.

The application has to be located together with folders 
"VirtualAudioDevice32" and "VirtualAudioDevice64" with 32 bit and 64 bit
versions of the dll as DirectShow Filter ".ax" files.

### Binaries

Binaries of the virtual audio capture device are included to illustrate the function of the program.
The dll files have been renamed from to ".ax". Original dll files can be downloaded from the author's repository.

### Copyright

VirtualAudioRegister - Copyright (C) 2023 Lynn Jarvis [http://spout.zeal.co/](http://spout.zeal.co/)\
VirtualAudioRegister is licenced using the GNU General Public License.


