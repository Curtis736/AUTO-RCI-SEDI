



//////////// PROJECT TREE ORGANIZATION ////////////


- Interface.py

    The main file of the project. Contains the necessary code to generate the main window, and call the init of every tab.


- Settings.py

    File that provides method to access application configs stored in .json files, in the config folder.
    Aims to make function loading and saving as foolproof and centralized as possible.
    When requested, config values are loaded if possible, or created.
    Many things such as the app's fields last values and paths are stored there.


- DefaultSettings.py

    Some configs should always exist with a valid value, such as the theme color.
    Provides method to complete the loaded config with any missing value


- Log.py

    Allows to output messages or errors to files, the standard output or callbacks.
    Always use this instead of print().


- PathsInterface.py

    Simple interface to select the three main paths necessary to document generation


- DocumentGenerator.py

    ...