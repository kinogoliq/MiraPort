# pyoxidizer.bzl

def make_dist():
    return default_python_distribution(
        name="proforma_app",
        python_config=python_config(
            executable_name="ProformaApp",
            use_lib_python=True,
        ),
        src_root=".",
        entry_module="main",
        resources=[
            Resource(
                path="icons",
                dest="icons",
                is_excluded=False,
            ),
            Resource(
                path="templates",
                dest="templates",
                is_excluded=False,
            ),
        ],
        macos_app=MacOSApp(
            app_name="ProformaApp",
            icon="icons/app_icon.icns",
        ),
    )
