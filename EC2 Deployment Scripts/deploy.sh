pkill -f acadedge

rm -f ../acadedge/acadedge-0.0.1-SNAPSHOT.jar

cp acadedge-0.0.1-SNAPSHOT.jar ../acadedge/

rm -f acadedge-0.0.1-SNAPSHOT.jar

cd ../acadedge/

sh ./start.sh
