#!/bin/bash

# Update the package list and install necessary dependencies
echo "Updating package list..."
sudo apt-get update

echo "Installing required packages..."
sudo apt-get install -y python3 python3-pip python3-venv unzip wget curl

# Create a virtual environment
echo "Creating virtual environment..."
python3 -m venv venv

# Activate the virtual environment
echo "Activating virtual environment..."
source venv/bin/activate

# Install Python packages
echo "Installing Python packages..."
pip install -r requirements.txt

# Ensure script is executable
echo "Making the schedule sync script executable..."
chmod +x schedule_sync.py

# Deactivate virtual environment
echo "Deactivating virtual environment..."
deactivate

echo "Setup complete. To run the script, activate the virtual environment using 'source venv/bin/activate' and run 'python schedule_sync.py'."
