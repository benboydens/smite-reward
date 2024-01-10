from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need fine tuning.
setup(
    executables=[
        Executable(
            script="smreward.py", 
            base="console",
            target_name="smreward",
            icon="logo.ico",
            shortcut_name="sm_reward_startup",
            shortcut_dir="StartupFolder"
        )
    ]
)