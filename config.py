from yaml import safe_load

CONFIG_PATH = 'data_storage/config.yaml'


def get_config():
    global CONFIG_PATH
    with open(CONFIG_PATH, "r") as stream:
        config = safe_load(stream)
        return config
