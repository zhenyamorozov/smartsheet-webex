# These exceptions may be raised at different stages of the scheduling proces


class ParameterStoreError(Exception):
    """Custom exception raised when there is a problem with Parameter Store"""
    pass

class SmartsheetInitError(Exception):
    """Custom exception raised when Smartsheet fails to initialize"""
    pass

class SmartsheetColumnMappingError(Exception):
    """Custom exception raised when Smartsheet columns could not be fully mapped to the schema"""
    pass

class WebexIntegrationInitError(Exception):
    """Custom exception raised when Webex Integration fails to initialize"""
    pass

class WebexBotInitError(Exception):
    """Custom exception raised when Webex Bot fails to initialize"""
    pass
