import csv
import platform
import time
from datetime import datetime

import psutil

CSV_COLUMNS = ("time", "cpu_%", "ram_%", "disk_%", "sent_MB", "recv_MB")


def sample_metrics(cpu_interval=1):
    cpu = psutil.cpu_percent(interval=cpu_interval)
    ram = psutil.virtual_memory().percent
    disk = psutil.disk_usage("/").percent
    net = psutil.net_io_counters()

    return {
        "time": datetime.now().strftime("%H:%M:%S"),
        "cpu_%": cpu,
        "ram_%": ram,
        "disk_%": disk,
        "sent_MB": round(net.bytes_sent / 1024 / 1024, 2),
        "recv_MB": round(net.bytes_recv / 1024 / 1024, 2),
    }


def main():
    print("System:", platform.system(), platform.release())
    print("CPU cores:", psutil.cpu_count())

    seconds = 5
    readings = 6

    with open("metrics.csv", "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(CSV_COLUMNS)

        for _ in range(readings):
            reading = sample_metrics()
            writer.writerow(reading[column] for column in CSV_COLUMNS)
            print(
                reading["time"],
                "cpu",
                reading["cpu_%"],
                "ram",
                reading["ram_%"],
                "disk",
                reading["disk_%"],
            )
            time.sleep(seconds)

    print("saved to metrics.csv")


if __name__ == "__main__":
    main()
