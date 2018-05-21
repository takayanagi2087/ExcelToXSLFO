#!/bin/sh
SCRIPT_DIR=`dirname "${0}"`
if [ "$JAVA_HOME" == "" ]; then
	JAVACMD=java
else
	JAVACMD=$JAVA_HOME/bin/java
fi
$JAVACMD $JAVAOPTS -cp $SCRIPT_DIR/excel2xslfo.jar:$SCRIPT_DIR/lib/* exeltoxslfo.ExcelToXSLFO $*
