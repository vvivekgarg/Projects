cd "C:\VivekData\Vivek_Personal\Learning\acadedge\Important"

rm acadedge-0.0.1-SNAPSHOT.jar

cp "C:\VivekData\Vivek_Personal\Learning\repo\acadedge\target\acadedge-0.0.1-SNAPSHOT.jar" ./

scp -i acadedge-app-vm-key-pair.pem -r acadedge-0.0.1-SNAPSHOT.jar ubuntu@ec2-13-233-152-182.ap-south-1.compute.amazonaws.com:~/
