# UltraPad

UltraPad is a modern and unofficial replacement for WordPad, written in C#. It is free and open source, and aims to provide almost the same functionality as its predecessor, with a fresh and updated look.

![UltraPad Screenshot](https://github.com/Lixkote/WordPad11/blob/main/preview_dark.png)
![UltraPad Screenshot](https://github.com/Lixkote/WordPad11/blob/main/preview_light.png)

## Features
- Standard WordPad features, such as basic text operations and inserting various multimedia items.
- Modern Windows 11-style design, based on the new paint redesign.

## Credits
 - [Jd](https://github.com/Jd-1206) for the initial idea to recreate WordPad in Windows 11 style.

## Compilation
To compile UltraPad, you will need the following prerequisites:
- A computer running Windows 11, build 22000 or newer.
- The latest version of [Visual Studio](https://developer.microsoft.com/en-us/windows/downloads) (the free community edition is sufficient).
  - The "Universal Windows Platform Development" workload installed.
  - The latest Windows 11 SDK and Windows 10 SDKs installed.


- The source code of UltraPad, which you can get by cloning this repository:
    ```
    git clone https://github.com/lkt27/UltraPad
    ```

- Open [src\WordPad.sln](/src/WordPad.sln) in Visual Studio to build and run the UltraPad app.
- If some icons appear as empty boxes, please install the font provided here: https://github.com/lkt27/UltraPad/blob/Preview1/WordPad/Assets/WordPadIcons.ttf
- Now you can press the compile button and the app should be ran in debug mode.

## Donate
If you find this project useful, or you like what i am doing and my work, please consider buying me a cofee to support the development through the kofi link button at the bottom of this message, so that i can continue to keep fixing bugs, and implement new features like ODT/DOCX support. Much thanks for any support :)
- [Buy me a cofee](https://ko-fi.com/lixkote)

## Contributing
If you want to contribute to UltraPad, please send me an email to lkt27@gmail.com.

## Projects used
 - [TowPad](https://github.com/itsWindows11/TowPad) for some of the mechanics.
 - [Appxinstaller](https://github.com/aL3891/AppxInstaller/tree/master) for converting appx into msi.
  
UltraPad is licensed under the [MIT License](./LICENSE).
