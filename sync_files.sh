#!/bin/bash

#.projectids is a file present only in the local environment
source .projectids

cd src/

echo "Pulling files from develop..."
echo $DEVELOP > ../.clasp.json
clasp pull

echo "Updating files on template file"
echo $TEMPLATE > ../.clasp.json
clasp push -f

echo "Updating files on operational file"
echo $OPERATIONAL > ../.clasp.json
clasp push -f

# restore clasp file back to develop
echo $DEVELOP > ../.clasp.json
echo "Success."
