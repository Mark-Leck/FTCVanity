HOW TO IMPORT THE GENERATED PRIVATE KEYS:

To import the generarted private key, simply start the QT client as usual and then navigate to:

Help >> Debug Window >> Console

And enter the following command:

importprivkey "the_VERY_long_generated_private_key_here" 

Press enter and once it has been verified with the blockchain you will then have your own personal Vanity address!

NOTE: The wallet.dat needs to be unencrypted in order to import an address, if your wallet is currently encrypted you can temporarily decrypt the wallet by entering the following command in the console:

walletpassphrase YourPassword Seconds

Replace 'Seconds' with the number of seconds you wish to unlock your wallet for (eg: 60 for 60 seconds, etc, etc), continue with importprivkey as usual.

For feedback or questions: http://forum.feathercoin.com/index.php?topic=489.0
