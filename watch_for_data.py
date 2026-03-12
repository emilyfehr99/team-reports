
import time
import shutil
import os
import subprocess
import sys

def watch_and_download():
    repo_dir = "/Users/emilyfehr8/CascadeProjects/team-reports"
    csv_path = os.path.join(repo_dir, "data/players_2025_26.csv")
    desktop_path = os.path.expanduser("~/Desktop/players_2025_26.csv")
    
    print(f"Monitoring {repo_dir} for updates...")
    print("Will automatically copy to Desktop when new data arrives.")

    # Get initial file size/mod time if it exists
    last_size = os.path.getsize(csv_path) if os.path.exists(csv_path) else 0
    
    while True:
        try:
            # Fetch and pull latest changes
            subprocess.run(["git", "pull", "origin", "main"], cwd=repo_dir, capture_output=True)
            
            if os.path.exists(csv_path):
                current_size = os.path.getsize(csv_path)
                
                # If file has grown significantly (indicating data arrival)
                # A full file will be >22KB (likely ~1MB+)
                if current_size > 50000 and current_size != last_size:
                    print(f"New data detected! Size: {current_size} bytes")
                    shutil.copy2(csv_path, desktop_path)
                    print(f"✅ SUCCESS: Copied to {desktop_path}")
                    print("You can open the file now.")
                    break
                elif current_size != last_size:
                     print(f"File updated but small ({current_size} bytes). waiting...")
                     last_size = current_size
            
            # Wait 30 seconds before checking again
            time.sleep(30)
            
        except Exception as e:
            print(f"Error checking for updates: {e}")
            time.sleep(30)

if __name__ == "__main__":
    watch_and_download()
