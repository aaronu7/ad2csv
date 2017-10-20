# ad2dt

This module was developed as part of a User Syncronization automation process which updates Active Directory based on the operational data from human resources. In production this module is configured as a sceduled console application and a triggerable web-service that will write the DataTable out to a CSV in a temporal repository (using csvplus-read-write https://github.com/aaronu7/csvplus-read-write).



Potential Upgrades:
- Extract Group memberships
- Extract additional user attributes