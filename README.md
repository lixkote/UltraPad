# UltraPad

The UltraPad app is a modern unofficial replacement for Windows WordPad. 
It is free and open source, aims to provide the same functionality as its predecessor, with a fresh and updated look, and some new useful features here and there.

![UltraPad Screenshot in light mode](https://raw.githubusercontent.com/lixkote/ultrapad/refs/heads/main/Previews/light.png)
![UltraPad Screenshot in dark mode](https://raw.githubusercontent.com/lixkote/ultrapad/refs/heads/main/Previews/dark.png)

## Features
- Standard WordPad features, such as rich text editing and inserting various multimedia items.
- Text margin ruler
- As lightweight as possible
- Modern Windows 11-style design, based on the new paint redesign.

## Compilation
To compile UltraPad, you will need the following prerequisites:
- A computer running Windows 11 24H2 or newer.
- The 2026 insider version of Visual Studio (the free community edition is sufficient).
  - The "WinUI app development" workload installed.
  - Windows 11 26000 SDK.


- The source code of UltraPad, which you can get by cloning this repository:
    ```
    git clone https://github.com/lixkote/UltraPad
    ```

- Open [src\WordPad.sln](/src/WordPad.sln) in Visual Studio to build and run the UltraPad app.
- Now you can press the compile button and the app should be ran in debug mode.

## Donate
If you find this project useful, or you like what i am doing and my work, please consider buying me a cofee to support the development through the kofi link button at the bottom of this message, so that i can continue to keep fixing bugs, and implement new features like ODT/DOCX support. Much thanks for any support :D
- [Buy me a coffee](https://ko-fi.com/lixkote)

## Contributing
If you want to contribute to UltraPad, please create a pull request.

## Projects used
 - [TowPad](https://github.com/itsWindows11/TowPad) for some of the mechanics.
 - [Appxinstaller](https://github.com/aL3891/AppxInstaller/tree/master) for converting appx into msi.
  
UltraPad is licensed under the [GPL-2.0 License](./LICENSE).
