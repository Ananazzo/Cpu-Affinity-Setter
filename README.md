# Cpu-Affinity-Setter
Sets what cores a program uses with a modern Intel CPU

# Usage:
   python cpu_affinity_service_groups.py install
   python cpu_affinity_service_groups.py start
   python cpu_affinity_service_groups.py stop
   python cpu_affinity_service_groups.py remove

# Requirements:
  pip install pywin32 psutil

# Notes:
 - Config file cpu_affinity_config.json in same folder.
 - Config uses "group": "P"|"E"|"ALL" (case-insensitive). Legacy list form [0,1] still accepted.
 - Smart Unlock: expands to ALL logical CPUs when assigned group is saturated.
 - Service must be installed/started with Administrator privileges.
