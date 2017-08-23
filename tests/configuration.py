import sys

APP_NAME = 'TestApp'

GROUNDWORK_LOGGING = {
    'version': 1,
    'disable_existing_loggers': False,
    'formatters': {
        'standard': {
            'format': "%(asctime)s - %(levelname)-5s - %(message)s"
        },
        'debug': {
            'format': "%(asctime)s [%(levelname)s] [%(name)s.%(funcName)s:%(lineno)s] %(message)s"
        }
    },
    'handlers': {
        'console_stdout': {
            'formatter': 'debug',
            'class': 'logging.StreamHandler',
            'stream': sys.stdout,
            'level': 'DEBUG'
        }
    },
    'loggers': {
        'EmptyPlugin': {
            'handlers': ['console_stdout'],
            'level': 'DEBUG',
            'propagate': True
        },
        'groundwork': {
            'handlers': ['console_stdout'],
            'level': 'INFO',
            'propagate': True
        }
    }
}
