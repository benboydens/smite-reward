# Smite reward script

Python script for collecting daily rewards in Smite. It will locate your steam installation and run Smite game. Once in the game inputs are given and the reward is collected. Afterwards the script shuts down the computer (this can be turned off).

## Dependencies

-  psutil, toml
-  win32api
-  cx_freeze

## TODO

- [ ] let script wait for steam updates
- [ ] detect if user started pc (to avoid anoying pop up when pc starts)
- [x] shut down computer when done
- [x] find steam executable
