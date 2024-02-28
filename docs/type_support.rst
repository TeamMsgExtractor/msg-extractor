Type Support
============

This page lists how much support a certain object has. There are three classifications: Open, meaning that the file can be opened but not saved at all, Partial, meaning that the file uses default saving characteristics and has not been fully implemented, and Full, meaning that the class has been completely written. A class with incomplete properties may still be listed as Full if the saving capabilities are done. In addition, the version for each milestone of a class is listed.

The first column is the internal class type that is used to figure out the value for the second column, the extract-msg class that is used to handle it. If the Class Name column is blank, look to the next one down for details, as it shares a class. Message class types ending with `.*` must start with the string but may have anything after it. If there is no more specialized version later in the table, anything starting with that will be handled by the specified class.

For things added before 0.29.0, 0.29.0 is listed as when they were added, as that is the oldest version officially supported.

Saving before version 0.35.0 may work for some things, but is not considered complete.

.. csv-table::
    :file: type-support.csv
