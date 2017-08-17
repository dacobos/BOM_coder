#! /bin/bash
# Create virtualenv

BASEDIR=`pwd`

# Run this for the first time
if [ ! -d "$BASEDIR/environment" ]; then
    echo "Installing..."
    virtualenv -q $BASEDIR/environment --no-site-packages
    echo "Virtualenv created."
    source $BASEDIR/environment/bin/activate
    cd $BASEDIR
    export PYTHONPATH=.
    echo "Installing dependencies..."
    pip install xlrd
    pip install xlutils
    pip install openpyxl
    echo "Dependencies installed."
fi

# Run this every time
source $BASEDIR/environment/bin/activate
cd $BASEDIR
export PYTHONPATH=.
echo "Running python script."
exec python $BASEDIR/__init__.py $1
