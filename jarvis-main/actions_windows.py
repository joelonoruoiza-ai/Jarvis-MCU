import subprocess
import webbrowser

class WindowsActions:
    @staticmethod
    def open_terminal():
        """Opens a PowerShell terminal."""
        subprocess.Popen(['powershell.exe'])

    @staticmethod
    def open_browser(url):
        """Opens a web browser to the specified URL."""
        webbrowser.open(url)

    @staticmethod
    def get_current_time():
        """Returns the current UTC date and time in formatted string."""
        from datetime import datetime
        return datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')

if __name__ == '__main__':
    actions = WindowsActions()
    actions.open_terminal()
    actions.open_browser('http://example.com')
    print(actions.get_current_time())