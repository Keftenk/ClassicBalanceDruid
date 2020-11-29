<p align='center'>
<a href="https://github.com/Keftenk/ClassicBalanceDruid/"><img src="Images/moonkin_banner.png"></a>
</p>

<p align='center'>
<a href="https://kmmiles.gitlab.io/moonkin-calc/"><img height="30" src="Images/moonfireicon.png"></a>&nbsp;&nbsp;
<a href="https://www.paypal.com/paypalme2/keftenk?locale.x=en_US"><img height="30" src="Images/paypalicon.png"></a>&nbsp;&nbsp;
</p>


# "By the great winds, I come." Classic Balance Druid Spreadsheet v1.8

One of the most advanced World of WarCraft Classic theorycraft resources for Balance Druid:

<img align="right" src="Images/moonkin_side.png" height="280">

- Optimize your gear selection through various presets
- Determine your optimal trinket rotation through the trinket optimizer
- Emulate specific encounter scenarios (duration, buffs, debuffs, etc.)
- Gauge your Moonkin Aura DPS contribution to your party and raid
- Visualize how stats from gear, buffs, world buffs, and more adjusts your critical stat weights
- Save and publish your gear sets for later use or comparison
- Utilize high-risk, high-reward tables to maximize your gameplay
- Add or remove gear which you've collected to optimize any preset specifically for <b><i>YOUR</i></b> in game scenario
- And more...!

This spreadsheet is a resource which is primarily for Excel to utilize the power and functionality of VBA macros. <b>You will not be able to use this tool in Google Docs!</b>

  ---
  
<p align='center'>
<a href="#About"><b>About</b></a>&nbsp;|&nbsp;
<a href="#Spreadsheet"><b>Spreadsheet</b></a>&nbsp;|&nbsp;
<a href="https://www.paypal.com/paypalme2/keftenk?locale.x=en_US"><b>Target Spell Resistance</b></a>&nbsp;|&nbsp;
<a href="https://www.paypal.com/paypalme2/keftenk?locale.x=en_US"><b>Web Application</b></a>&nbsp;|&nbsp;
<a href="https://www.paypal.com/paypalme2/keftenk?locale.x=en_US"><b>FAQ</b></a>&nbsp;|&nbsp;
<a href="https://www.paypal.com/paypalme2/keftenk?locale.x=en_US"><b>Author Notes</b></a>&nbsp;|&nbsp;
<a href="https://www.paypal.com/paypalme2/keftenk?locale.x=en_US"><b>Donation</b></a>&nbsp;|&nbsp;
<a href="https://www.paypal.com/paypalme2/keftenk?locale.x=en_US"><b>Credit</b></a>&nbsp;|&nbsp;
</p> 

  ---

# <a href id="#About"></a>About

This project was developed not out of the desire to play the Moonkin myself, but to justifiably refute accusations within the community as to how Balance Druid performs in a raid environment and if there was any validity to the claims players would make or if everything was hearsay. I went in with a unbiased opinion of the specs performance and detailed the process backed by mathematics. I can say with enough assurance that I can speak to the Balance Druid better than most. I hope you all can appreciate and enjoy the resource.

Balance Druid's stereotype which has carried through to today has been the notion that they run out of mana exceedingly fast, being coined "OOMkin". This is a verifiably false statement with the capability of being able to last nearly 5-minutes in the current state of the game. In hindsight Smite Priest, Shadow Priest, as well as Elemental Shaman have more mana management opportunities.

Moonkin as well as the Druid talent tree revamp came very late in the progression of Vanilla. The understanding and knowledge of the specialization couldn't be fully realized even partially through the evolution of the game due to these changes. In tandem with a bottom tier played class made for stand out research and data building incredibly sparse and inconsistent.


  ---

# <a href id="#Spreadsheet"></a>Spreadsheet

<img align="left" src="Images/optimize.png" height="280">

By utilizing VBA macros you may optimize your gear down to exact value of a stat such as: Spell Hit, Spell Crit, Spell Damage, and more. This can be achieved by selecting your desired parameters in the <b>Character Selection</b> row and any additional <b>Settings</b> such as talents, buffs, and debuffs. Once selected you may click <b>Final Optimal Gear Set</b> and if you desire an even more exact optimization then click <b>Exact Optimization</b> at the bottom of the page.

&nbsp;&nbsp;

The Score function is the culmination of the stat weights via each individual item which has been selected. All of these parameters can be saved and called upon at a later load of the resource.

The spreadsheet also comes equipped with a rough estimation of the DPS that you may be able to produce with the given parameters. Tthis is not a perfect simulation of damage and does have a inherent margin of error which comes with it. Additionally, Zephan from the Classic Warlock community designed a <b>Moonkin Aura</b> tool which generates a rough estimation of the DPS contribition it brings to the party and raid.

<sub><sup><i>NOTE: Currently Ignite is not supported and I'm looking for someone able to reconstruct the tool appropriately for the aura with Ignite included.</i></sub></sup>

# Trinket Optimization

This feature of the tool is extremely powerful, but is easily the most miss appropriated function of the entire sheet. As a summary the optimization done in this worksheet will allow you to optimize for specific duration fights, such as a raid boss, or even optimize a trinket rotation for a given duration for trash monsters. I will detail the basic function and steps in order to successfully utilize the tool:

- Depending on what <b>Phase</b> your parameter is set to on the <b>Character</b> sheet you will be able to select or deselect any trinkets you wish to optimize in your test
- If optimizing for a boss make sure that <b>Boss?</b> is selected to <b>Yes</b> and a duration is input in terms of <b>seconds</b>
- Click <b>Generate Item Assignments</b> and then <b>Optimization</b>
- Scroll over to the right of this page and you will see the breakdown of the trinkets it selected as optimal
- If you wish to optimize the trinket selection with the rest of your gear click over to the <b>Character</b> sheet and make sure the <b>Optimize Trinket</b> parameter in the <b>Character Selection</b> column is selected as <b>Yes</b>
- Click <b>Find Optimal Gear Set</b> and then <b>Exact Optimization</b> at the bottom of the page if desired

  ---
