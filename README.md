# ad2dt

This module was developed as part of a User Syncronization automation process which updates Active Directory based on the operational data from human resources. In production this module is configured as a sceduled console application and a triggerable web-service that will write the DataTable out to a CSV in a temporal repository (using csvplus-read-write https://github.com/aaronu7/csvplus-read-write).

Project Notes:
- Developed in VS2015 using the nunit package for unit testing.
- Download the "NUnit 3 Test Adapter" extension to view example tests in the Test Explorer
- Although a trivial application implements the module, examples of usage are best exemplified through the unit tests.

Potential Upgrades:
- Extract Group memberships
- Extract additional user attributes