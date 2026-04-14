import os
import sys

# Platform-aware action routing
def route_action(action_name):
    os_type = sys.platform
    
    if os_type == 'win32':
        import windows_specific_module as action_module
    elif os_type == 'darwin':
        import mac_specific_module as action_module
    elif os_type.startswith('linux'):
        import linux_specific_module as action_module
    else:
        raise ImportError(f'Unsupported operating system: {os_type}')
    
    return action_module.perform_action(action_name)