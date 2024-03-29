================================
Smartsheet and Webex Automation
================================
.. image:: https://static.production.devnetcloud.com/codeexchange/assets/images/devnet-published.svg
    :alt: published
    :target: https://developer.cisco.com/codeexchange/github/repo/zhenyamorozov/smartsheet-webex

*Automatically create webinars in Webex Webinar based on information in Smartsheet*


It is easy to use:

Collaborate with your team on webinar planning. When ready for creation, check **Create=yes**

.. image:: docs/images/smartsheet-screenshot.gif
    :width: 1500
    :alt: Smartsheet screenshot

Schedule all webinars with one bot command

.. image:: docs/images/bot-screenshot.gif
    :width: 854
    :alt: Webex bot screenshot

Webinars are created

.. image:: docs/images/smartsheet-done-screenshot.gif
    :width: 1500
    :alt: Smartsheet done screenshot

If need to change title, description, or reshedule, run the bot command again, or set it to run on a schedule.


Features
--------
This automation ties together three different services: Smartsheet, Webex Meetings/Webinars and Webex Messaging bot. It helps a lot if you are running many webinars, especially in series, especially with multiple people collaborating.

This automation supports:

- Create and update Webex Webinars based on information in Smartsheet
- Reports status via bot to a Webex space
- Control with Webex bot adaptive cards
- Creation can be triggered by bot command or by schedule
- Customizable webinar parameters
- Attendee link, host key and registrant count updated into Smartsheet


How it works
------------

- Collect all webinar information in a smartsheet, one webinar per row. Include details like webinar title, description, date and time, hosts, panelists etc. The smartsheet can be shared by multiple people for teamwork.
- Check out individual webinars for creation by changing the ``Create`` field to ``yes``. Save the smartsheet.
- Mention the @bot in the Webex room and click ``Schedule now`` button.
- The scheduling will be triggered and the bot will report back after some seconds (or minutes, depending on your amount of webinars).


Get Started
-----------

This automation requires a few things to be set up. Look for details in `Get Started <docs/get_started.rst>`_


Contribute
----------

Feel free to fork and improve.


Support
-------

This automation is offered as-is.
