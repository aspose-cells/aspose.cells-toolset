
import matplotlib.pyplot as plt
import matplotlib.axes 
import numpy as np
import pandas as pd
import os
import sys
import unittest
import warnings
import re


class TestCellsToolset( unittest.TestCase):
    
    def setUp(self):
       
        warnings.simplefilter('ignore', ResourceWarning)
    
    def tearDown(self):
        pass
    def test_simply(self):
        df = pd.DataFrame([['a',1,2],['b',2,5],['c',3,3]],columns=['x','y','z'])
        print(df)
        df = pd.DataFrame(np.zeros((3,3)),columns=['x','y','z'])
        print(df)
        # Creating the Series 
        Cut = pd.Series(['New York', 'Chicago', 'Toronto', 'Lisbon', 'Rio', 'Moscow']) 
        Cut.astype("category",categor =['New York', 'Chicago', 'Toronto', 'Lisbon', 'Rio', 'Moscow'] ,order=True)
        print(Cut)
        
        # Create the Datetime Index 
        # didx = pd.DatetimeIndex(start ='2014-08-01 10:00', freq ='W',periods = 6, tz = 'Europe/Berlin')  
        # Cut = ["Fair" , "Good" , "Very Good" , "Premium" , "Ideal"]
        # Cut_Facor = pd.Categorical(Cut,categories=Cut ,ordered=True)
        # print(Cut_Facor)
        # set the index 
        # sr.index = didx         
        # print(didx)
        # fig, ax = plt.subplots()  # Create a figure containing a single axes.
        # ax.set_title('Ticks seem out of order / misplaced')
        # axf = ax.plot([1, 2, 3, 4], [1, 4, 2, 3])  # Plot some data on the axes.
        # print(axf)
        # plt.show()
        pass
        
    # def test_data_plot(self):
    #     fig, ax = plt.subplots(1, 2, layout='constrained', figsize=(6, 2))

    #     ax[0].set_title('Ticks seem out of order / misplaced')
    #     x = ['5', '20', '1', '9']  # strings
    #     y = [5, 20, 1, 9]
    #     ax[0].plot(x, y, 'd')
    #     ax[0].tick_params(axis='x', labelcolor='red', labelsize=14)

    #     ax[1].set_title('Many ticks')
    #     x = [str(xx) for xx in np.arange(100)]  # strings
    #     y = np.arange(100)
    #     ax[1].plot(x, y)
    #     ax[1].tick_params(axis='x', labelcolor='red', labelsize=14)
    #     print(ax)
    #     print(ax[0].title)
    #     print(ax[1].title)
        
    #     pass
    
if __name__ == '__main__':
    unittest.main()
        