# -*- coding: utf-8 -*-
"""
Created on Fri Dec 17 12:54:05 2021

@author: jacob.sterling
"""
from zcrmsdk.src.com.zoho.crm.api.user_signature import UserSignature
from zcrmsdk.src.com.zoho.crm.api.dc import USDataCenter
from zcrmsdk.src.com.zoho.api.authenticator.store import FileStore
from zcrmsdk.src.com.zoho.api.logger import Logger
from zcrmsdk.src.com.zoho.crm.api.initializer import Initializer
from zcrmsdk.src.com.zoho.api.authenticator.oauth_token import OAuthToken, TokenType
from zcrmsdk.src.com.zoho.crm.api.sdk_config import SDKConfig
from zcrmsdk.src.com.zoho.crm.api.query import response_wrapper

class SDKInitializer(object):
    @staticmethod
    def initialize():
        logger = Logger.get_instance(level=Logger.Levels.INFO, 
                                     file_path= r'C:\Users\jacob.sterling\OneDrive - advance.online\Documents\PythonSDK\python_sdk_log.log')
        user = UserSignature(email='jacob.sterling@advance.online')
        environment = USDataCenter.PRODUCTION()
        """
        scope = AAAServer.profile.read,ZohoCRM.Modules.ALL,ZohoCRM.Bulk.READ,ZohoCRM.org.all
        """
        token = OAuthToken(client_id='1000.4VZAY9ESY6Z2ZSSDRN932MAA06WSNV', 
                           client_secret='ebd361647e558d88795e40b79cc95ce78272bf2162',
                           token='1000.b67b00bbb84ed3d2cf0f6e6d6a7f726e.a503419fe2779902733d9bce58e276e1',
                           token_type= TokenType.GRANT,
                           redirect_url='redirectURL')
        
        store = FileStore(file_path=r'C:\Users\jacob.sterling\OneDrive - advance.online\Documents\PythonSDK/python_sdk_tokens.txt')
        config = SDKConfig(auto_refresh_fields=True, pick_list_validation=False)
        resource_path = r'C:\Users\jacob.sterling\OneDrive - advance.online\Documents\PythonSDK\python-app'
        
        Initializer.initialize(user=user, environment=environment, store=store, token=token, sdk_config=config, resource_path=resource_path, logger=logger)
        
SDKInitializer.initialize()

response_wrapper.ResponseWrapper.get_data('Contacts')