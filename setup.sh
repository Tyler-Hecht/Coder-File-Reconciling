python3 clearer.py 
if [ $? -ne 0 ]; then exit 0; fi 
python3 setup1.py 
if [ $? -ne 0 ]; then exit 0; fi 
cd DatavyuToSupercoder 
java -jar DatavyuToSupercoder.jar 
cd .. 
python3 setup2.py 
if [ $? -ne 0 ]; then exit 0; fi 