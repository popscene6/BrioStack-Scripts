import pathlib
import subprocess
import os

renewal = pathlib.Path(r'Y:\BrioStack\Script Uploads\Renewal Letters.csv')
retreat = pathlib.Path(r'Y:\BrioStack\Script Uploads\Retreat Letters.csv')
final = pathlib.Path(r'Y:\BrioStack\Script Uploads\Final Letters.csv')

if renewal.exists():
    print("Renewal")
    import renewal_letters

if retreat.exists():
    print("Retreat")    
    import retreat_letters  

if final.exists():
    print("Final")
    import finals

else: 
    print ("None!")
   

import closeout

quit()
