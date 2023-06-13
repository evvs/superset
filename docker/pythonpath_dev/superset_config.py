# Licensed to the Apache Software Foundation (ASF) under one
# or more contributor license agreements.  See the NOTICE file
# distributed with this work for additional information
# regarding copyright ownership.  The ASF licenses this file
# to you under the Apache License, Version 2.0 (the
# "License"); you may not use this file except in compliance
# with the License.  You may obtain a copy of the License at
#
#   http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing,
# software distributed under the License is distributed on an
# "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
# KIND, either express or implied.  See the License for the
# specific language governing permissions and limitations
# under the License.
#
# This file is included in the final Docker image and SHOULD be overridden when
# deploying the image to prod. Settings configured here are intended for use in local
# development environments. Also note that superset_config_docker.py is imported
# as a final step as a means to override "defaults" configured here
#
import math
import logging
import os
from typing import Optional
 
from cachelib.file import FileSystemCache
from celery.schedules import crontab
from datetime import timedelta
 
logger = logging.getLogger()
 
 
def get_env_variable(var_name: str, default: Optional[str] = None) -> str:
    """Get the environment variable or raise exception."""
    try:
        return os.environ[var_name]
    except KeyError:
        if default is not None:
            return default
        else:
            error_msg = "The environment variable {} was missing, abort...".format(
                var_name
            )
            raise OSError(error_msg)
 
 
DATABASE_DIALECT = get_env_variable("DATABASE_DIALECT")
DATABASE_USER = get_env_variable("DATABASE_USER")
DATABASE_PASSWORD = get_env_variable("DATABASE_PASSWORD")
DATABASE_HOST = get_env_variable("DATABASE_HOST")
DATABASE_PORT = get_env_variable("DATABASE_PORT")
DATABASE_DB = get_env_variable("DATABASE_DB")
 
# The SQLAlchemy connection string.
SQLALCHEMY_DATABASE_URI = "{}://{}:{}@{}:{}/{}".format(
    DATABASE_DIALECT,
    DATABASE_USER,
    DATABASE_PASSWORD,
    DATABASE_HOST,
    DATABASE_PORT,
    DATABASE_DB,
)
 
REDIS_HOST = get_env_variable("REDIS_HOST")
REDIS_PORT = get_env_variable("REDIS_PORT")
REDIS_CELERY_DB = get_env_variable("REDIS_CELERY_DB", "0")
REDIS_RESULTS_DB = get_env_variable("REDIS_RESULTS_DB", "1")
 
RESULTS_BACKEND = FileSystemCache("/app/superset_home/sqllab")
 
CACHE_CONFIG = {
    "CACHE_TYPE": "RedisCache",
    "CACHE_DEFAULT_TIMEOUT": int(timedelta(minutes=1).total_seconds()),
    "CACHE_KEY_PREFIX": "superset_",
    "CACHE_REDIS_HOST": REDIS_HOST,
    "CACHE_REDIS_PORT": REDIS_PORT,
    "CACHE_REDIS_DB": REDIS_RESULTS_DB,
}
DATA_CACHE_CONFIG = CACHE_CONFIG
 
 
class CeleryConfig:
    broker_url = f"redis://{REDIS_HOST}:{REDIS_PORT}/{REDIS_CELERY_DB}"
    imports = ("superset.sql_lab",)
    result_backend = f"redis://{REDIS_HOST}:{REDIS_PORT}/{REDIS_RESULTS_DB}"
    worker_prefetch_multiplier = 1
    task_acks_late = False
    beat_schedule = {
        "reports.scheduler": {
            "task": "reports.scheduler",
            "schedule": crontab(minute="*", hour="*"),
        },
        "reports.prune_log": {
            "task": "reports.prune_log",
            "schedule": crontab(minute="10", hour="0"),
        },
    }
 
 
CELERY_CONFIG = CeleryConfig
 
FEATURE_FLAGS = {"ALERT_REPORTS": True}
ALERT_REPORTS_NOTIFICATION_DRY_RUN = True
WEBDRIVER_BASEURL = "http://superset:8088/"
# The base URL for the email report hyperlinks.
WEBDRIVER_BASEURL_USER_FRIENDLY = WEBDRIVER_BASEURL
 
SQLLAB_CTAS_NO_LIMIT = True
 
#
# Optionally import superset_config_docker.py (which will have been included on
# the PYTHONPATH) in order to allow for local settings to be overridden
#
 
#manzana_change_config_start
 
SQLALCHEMY_POOL_SIZE = 45
 
SQLALCHEMY_MAX_OVERFLOW = 30
 
SQLALCHEMY_POOL_TIMEOUT = 180
 
SUPERSET_WEBSERVER_TIMEOUT = 1200
 
 
FILTER_STATE_CACHE_CONFIG = {
    "CACHE_TYPE": "SimpleCache",
    "CACHE_THRESHOLD": math.inf,
    "CACHE_DEFAULT_TIMEOUT": int(timedelta(minutes=10).total_seconds()),
}
 
EXPLORE_FORM_DATA_CACHE_CONFIG = {
    "CACHE_TYPE": "SimpleCache",
    "CACHE_THRESHOLD": math.inf,
    "CACHE_DEFAULT_TIMEOUT": int(timedelta(minutes=10).total_seconds()),
}
 
CSV_EXPORT = {
    "encoding": "utf-8-sig"
    ,"sep": ";"
    ,"decimal": ","
}
 
SCREENSHOT_LOCATE_WAIT = 300
SCREENSHOT_LOAD_WAIT = 600
 
EMAIL_REPORTS_USER = "admin"
THUMBNAIL_SELENIUM_USER = "admin"
 
# smtp server configuration
EMAIL_NOTIFICATIONS = True  # all the emails are sent using dryrun
SMTP_HOST = "" # test.test.ru
SMTP_STARTTLS = False
SMTP_SSL = True
SMTP_USER = "" # notifications
SMTP_PORT = 465 # SET PORT
SMTP_PASSWORD = "" # SET PASSWORD FOR SMTP
SMTP_MAIL_FROM = "" # test@ya.ru
 
# only ru language untill fix
BABEL_DEFAULT_LOCALE = "ru"
LANGUAGES = {
    'ru': {'flag': 'ru', 'name': 'Russian'},
}
 
SQL_MAX_ROW = 500000
 
DISPLAY_MAX_ROW = 500000
 
ROW_LIMIT  =  500000
 
SAMPLES_ROW_LIMIT  =  500000
 
FILTER_SELECT_ROW_LIMIT  =  500000
 
ALLOW_FULL_CSV_EXPORT = True
 
FEATURE_FLAGS = {"ALERT_REPORTS": True,
"ENABLE_TEMPLATE_PROCESSING": True,
"DASHBOARD_CROSS_FILTERS": True,}
 
ALERT_REPORTS_NOTIFICATION_DRY_RUN = False
 
#manzana_change_config_end
 
try:
    import superset_config_docker
    from superset_config_docker import *  # noqa
 
    logger.info(
        f"Loaded your Docker configuration at " f"[{superset_config_docker.__file__}]"
    )
except ImportError:
    logger.info("Using default Docker config...")