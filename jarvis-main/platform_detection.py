# platform_detection.py

import os
import sys
import platform

class PlatformDetection:
    def __init__(self):
        self.system = platform.system()
        self.platform_specific_modules = self.get_platform_specific_modules()

    def get_platform_specific_modules(self):
        if self.system == "Windows":
            return {"os": os, "sys": sys}
        elif self.system == "Linux":
            return {"os": os, "sys": sys}
        elif self.system == "Darwin":  # macOS
            return {"os": os, "sys": sys}
        else:
            raise NotImplementedError(f"Unsupported platform: {self.system}")

    def execute_command(self, command):
        if self.system == "Windows":
            os.system(command)
        else:
            os.system(command)

    def get_platform_info(self):
        return self.system, self.platform_specific_modules

# Example usage:
if __name__ == "__main__":
    detector = PlatformDetection()
    print(detector.get_platform_info())