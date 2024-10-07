#!/bin/bash

git checkout wb
git pull
git add .
git commit -m 'using commit script to working branch'
git push 

# git push --set-upstream origin wb