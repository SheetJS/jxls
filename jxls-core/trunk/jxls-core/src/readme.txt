jXLS - Export data to Excel using XLS template
===============================================
Online docs - http://jxls.sourceforge.net/

Installation
-------------
You can download the latest jXLS release from here:

    http://sourceforge.net/project/showfiles.php?group_id=141729

Put jxls.jar into classpath of your application

Dependencies
------------
jXLS requires next libraries to be on your classpath

Jakarta POI - great library to manipulate XLS files from pure Java

    http://jakarta.apache.org/poi/ 

Jakarta Commons BeanUtils - great library for dynamic defining and accessing bean properties.

   http://jakarta.apache.org/commons/beanutils/

Jakarta Commons Collections - great library for manipulating java collections.

   http://jakarta.apache.org/commons/collections/

Jakarta Commons JEXL - excellent library for Expression Language support.

   http://jakarta.apache.org/commons/jexl/

Jakarta Commons Digester - excellent library to create objects from XML

    http://jakarta.apache.org/commons/digester/

Jakarta Commons Logging - good logging library

    http://jakarta.apache.org/commons/logging/

Building from source
--------------------
Source code for jXLS can be downloaded from

    http://sourceforge.net/project/showfiles.php?group_id=141729

To build the project you need Apache Ant utility.
It can be found here :

  http://ant.apache.org/

For testing the project, you will also need JUnit :

  http://www.junit.org/

Put junit.jar in $ANT_HOME/lib.

To build project documentation from xdocs you have to install Apache Forrest project

    http://forrest.apache.org/

Once you have Ant properly installed, and the
build.properties file correctly reflects the location
of your required jars, you are ready to build and test.
The major targets are:

ant compile         - compile the code
ant test            - test using junit
ant jar             - create a jar file
ant javadoc         - build the javadoc
ant projectdoc      - build the project docs (using forrest)
ant dist            - create folders as per a distribution
ant release-binary  - build binary release distribution
ant release-src     - build source release distribution


Legal
-----------
This software is distributed under the terms of the FSF Lesser GNU Public License (see lgpl.txt).

This product includes software developed by the Apache Software Foundation (http://www.apache.org/).
