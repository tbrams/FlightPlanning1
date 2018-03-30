# FlightPlanning1
This is the first Google Sheet based planning tool I created while studying for my PPL license. This one is based on 
the Flight Plan used at Copenhagen Airtaxa and comes with support for a few of their TP-9 models. 

Everything is controlled from a Google Sheet with this code installed.

When opening the sheet the weather data is checked automatically. You will be warned if a newer METAR has been issued since you originally made the plan. 

<p align="center">
<img width="400" alt="Metar out of sync" src="https://user-images.githubusercontent.com/3058746/38125494-02a8e650-3414-11e8-86f8-61d935a09cc5.png">
</p>

Updating is done easily but might obviously change a lot of performance calculations.

<p align="center">
<img width="400" alt="Update Metar" src="https://user-images.githubusercontent.com/3058746/38125508-246f8d34-3414-11e8-9770-5c0a520c187b.png">
</p>
After that you might notice slight changes here and there due to changed weather


Start planning by using this menu
<p align="center">
<img width="400" alt="Start Planning" src="https://user-images.githubusercontent.com/3058746/38125518-33f36528-3414-11e8-8809-2643b5a4e793.png">
</p>

There are a couple of screens available. The first is for selecting airfields and runways for take off and landing.
<p align="center">
<img width="300" alt="Start and end selection" src="https://user-images.githubusercontent.com/3058746/38089849-02a5b8ba-338b-11e8-96d2-d1628a539581.png"></p>
</p>

The second is for adding WP and leg information after planning manually with pen and paper on a traditional map. Also this is where you add the alternate when you are done with the main route.

<p align="center">
<img width="300" alt="Route Planning" src="https://user-images.githubusercontent.com/3058746/38089909-25a55b36-338b-11e8-9507-455a7291b5e2.png">
</p>

The third is for Weight and Balance. Select a registration number from the drop down many and type in a value in each field. 
Use zero instead of blanks to make sure...

<p align="center">
<img width="300" alt="W/B 1" src="https://user-images.githubusercontent.com/3058746/38089935-37333cc4-338b-11e8-945e-3d60d5b24ad1.png">
</p>

After you are done, the envelope will be presented

<p align="center">
<img width="400" alt="W/B Envelope" src="https://user-images.githubusercontent.com/3058746/38089958-4b0a0674-338b-11e8-9307-2a8ac8b24d4f.png">
</p>

And finally you have the option to specify different conditions for Take Off and Landing. 

<p align="center">
<img width="300" alt="TO/LD Performance" src="https://user-images.githubusercontent.com/3058746/38125975-a625f6da-3417-11e8-947a-ba33d9d5b03a.png">
</p>

When all this is done, you should have three green on the spreadsheet indicating you have satisfied all conditions necessary for a safe trip from a planning point of view. (Obviously you still need to check NOTAMs etc on own...)

<p align="center">
<img width="300" alt="W/B Envelope" src="https://user-images.githubusercontent.com/3058746/38089982-5df4d70a-338b-11e8-84c4-7064fb519cb3.png">
</p>

The Plan is now available with all the calculated values for Ground Speed, TAS and all the different headings and timings in this table

<p align="center">
<img width="300" alt="Final Plan" src="https://user-images.githubusercontent.com/3058746/38125989-bd5f784e-3417-11e8-9e1f-84ef6d9315e2.png">
</p>


## TO BE WRITTEN
- Extending with your own airplane registration
- Managing the Mass Balance Envelope
- How to overcome issues parsing weather information
- Extending the Aerodrome database
- Details on performance and fuel calculations
- Details on calculating headings, TAS and GS 


