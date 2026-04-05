# main.py
import subprocess
import sys

def run_test():
    print("Running Exporter Test...")
    # Запускаем тестовый скрипт
    result = subprocess.run([sys.executable, "tests/test_exporter.py"])
    return result.returncode

if __name__ == "__main__":
    exit(run_test())