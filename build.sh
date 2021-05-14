#!/bin/bash

#Example user 
#. ./build.sh SendEmail 1.0.0
#Expect result SendEmail:10.0.0.zip


# Define Variable
export FUNCTION_NAME=$1
( 
    cd $FUNCTION_NAME &&

    # Buidling
    dotnet build  &&
    dotnet publish -c Release -o publish 

    # Zipping
    cd publish
    zip -r ../../$FUNCTION_NAME.zip *
)
