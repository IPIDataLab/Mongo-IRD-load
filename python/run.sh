#! /bin/bash
for i in ../data/Database*.xlsx; do  yes ""|ipython main.py "$i"; done
cat ../data/*json | perl -pi -e 's/[]][[]/,/;' > ../../sample-maps/data/IRD.json 
      