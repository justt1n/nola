echo "Running Script"
echo "Checking for update"
git reset --hard origin/master
git pull
echo "Update completed"
& myenv\Scripts\Activate.ps1
echo "Running Script..."
python script.py
echo "Complete"
