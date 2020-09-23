I wrote this component because I needed an effective way to manage Windows NT MTS servers from 
Windows 2000.

I wrote this in about three days, so I was a little tired during some of the coding. :)

The usage is pretty simple.

The DLL included (WOWMTSAdmin.dll) needs to be installed in it's own MTS package on the sever that you
want to administer.  It must have the ability to modify packages on that server.

This dll also needs to be registered on the client before the client will be able to talk to the server.

Once both of these steps are completed, the client can be run.

When new computers are added to the computer's list, the computer name should be used.

If you find any bugs or issues, please let me know at rakker91@hotmail.com.

Certain features are not yet implemented (such as Interfaces and Methods).  Also, no security features (roles) 
are implemented in this version.

Please let me know if you have any comments or questions.

Robert May

Updates to the original release:

Bug fix:  When adding new remote dll's to a package the dll name wouldn't be mapped to a UNC when the file was added from a
		remotely mapped drive.

BUG FIX:  If two remote machines had the same components loaded, an error would occur when browsing the treeview due to duplicate
		key names.
