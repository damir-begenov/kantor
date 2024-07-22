#!/bin/bash

# Maven clean install
echo "Running Maven clean install..."
 mvn clean install

# Check if Maven build was successful
if [ $? -ne 0 ]; then
    echo "Maven build failed. Deployment aborted."
    exit 1
fi

# Define SSH variables
REMOTE_HOST="10.202.20.92"  # Replace with your remote server's hostname or IP address
REMOTE_USER="sadmin"  # Replace with your remote server's username
REMOTE_DIR=""  # Replace with the directory path on the remote server where you want to transfer dossier-app/target/

# Transfer dossier-app/target/ to remote server using SSH
echo "Transferring dossier-app/target/ to remote server..."
scp -r dossier-app/target/dossier-app-0.0.1-SNAPSHOT.jar ${REMOTE_USER}@${REMOTE_HOST}:${REMOTE_DIR}

# Check if scp transfer was successful
if [ $? -ne 0 ]; then
    echo "SCP transfer failed. Deployment aborted."
    exit 1
fi

# SSH into remote server and run the JAR file
echo "Stopping previous Java process and running JAR file on remote server..."
ssh ${REMOTE_USER}@${REMOTE_HOST} << EOF
    pkill -f "java -jar dossier-app-0.0.1-SNAPSHOT.jar"  # Kill existing Java process by matching the command
    java -jar dossier-app-0.0.1-SNAPSHOT.jar &
EOF

# Check if SSH command was successful
if [ $? -ne 0 ]; then
    echo "Failed to run JAR file on remote server."
    exit 1
fi

echo "Deployment and execution successful."
exit 0