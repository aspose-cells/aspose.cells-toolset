import pandas as pd
import re
from aspose.cells.tables import ListObject
from aspose.cells import CellsHelper
from aspose.cells import CellValueType
from aspose.cells import Cells
from aspose.cells import Range
from aspose.cells import Worksheet
from aspose.cells.charts import Chart
from Aspose.Cells.Charts import ChartType

class DataConversion(object):
    
    def __init__(self):

        pass

    def listobject_to_dataframe( self , table : ListObject ) -> pd.DataFrame :
        return self.__get_dataframe(table.data_range.worksheet.cells ,table.start_row,table.start_column,table.end_row,table.end_column,table.show_header_row,table.show_totals )
        pass
    
    def range_to_dataframe( self , range_name : Range) -> pd.DataFrame:
        begin_row_index = range_name.first_row
        begin_column_index = range_name.first_column
        end_row_index = range_name.first_row + range_name.row_count -1
        end_column_index = range_name.first_column + range_name.column_count -1
        cells = range_name.worksheet.cells
        has_header = self.__has_table_header(cells,begin_row_index,begin_column_index,end_row_index,end_column_index)
        return self.__get_dataframe(cells ,begin_row_index,begin_column_index,end_row_index,end_column_index,has_header,False )    
        pass
    
    def worksheet_to_dataframe( self , worksheet : Worksheet) -> pd.DataFrame:
        cells = worksheet.cells
        begin_row_index = cells.min_data_row
        begin_column_index = cells.min_data_column
        end_row_index = cells.max_data_row 
        end_column_index = cells.max_data_column         
        has_header = self.__has_table_header(cells,begin_row_index,begin_column_index,end_row_index,end_column_index)
        return self.__get_dataframe(cells ,begin_row_index,begin_column_index,end_row_index,end_column_index,has_header,False )    
        pass

    def chart_to_plot( self, chart :Chart ):
        workbook =  chart.worksheet.workbook
        data = {}
        series = self.__parse_data_source( chart.n_series.category_data )
        cells = workbook.worksheets.get(series[0]).cells
        column_index  = series[2]
        column_data = []
        for row_index in range (series[1],series[3] +1):
            column_data.append(cells.get(row_index , column_index).value)
        
        xName = ""
        if cells.get(series[1] -1  , column_index).type == CellValueType.IS_NULL :
            xName = CellsHelper.column_index_to_name(series[1])
            
        else:
            xName = cells.get(series[1] -1  , column_index).value
            
        data[xName] = column_data
        yNames = []
        for index in range( 0, chart.n_series.count):
            values = self.__parse_data_source( chart.n_series.get(index).values)
            values_data = []
            for row_index in range (values[1],values[3] +1):
                values_data.append(cells.get(row_index , column_index).value)
            data[chart.n_series.get(index).display_name] = values_data
            yNames.append(chart.n_series.get(index).display_name)
        
        plot = pd.DataFrame(data).plot(x=xName,y=yNames,kind=self.__get_type(chart.type))
        return plot    
        pass
    
    def dataframe_to_listobject( self ,dataframe: pd.DataFrame, cells: Cells , first_row : int , first_column: int ) -> ListObject:
        cells_area = self.__dataframe_import_cells(dataframe ,cells,first_row ,first_column)
        return cells.first_cell.worksheet.list_objects.add(cells_area[0], cells_area[1], cells_area[0] + cells_area[2] - 1, cells_area[1] + cells_area[3] -1 ,True) 
        pass

    def dataframe_to_range(self,dataframe: pd.DataFrame , cells: Cells , first_row : int , first_column: int) -> Range:
        cells_area = self.__dataframe_import_cells(dataframe ,cells,first_row ,first_column)
        return cells.create_range(cells_area[0],cells_area[1],cells_area[2],cells_area[3])        
        pass
    
    def dataframe_to_worksheet(self,dataframe: pd.DataFrame , cells: Cells , first_row : int , first_column: int):
        cells_area = self.__dataframe_import_cells(dataframe ,cells,first_row ,first_column)
        return cells.first_cell.worksheet       
        pass        
    
    def __has_table_header(self, cells: Cells, begin_row_index :int, begin_column_index:int, end_row_index :int, end_column_index:int ):
        has_header = True
        for column_index in range(begin_column_index , end_column_index +1) :
            cell = cells.get(begin_row_index,column_index)
            if cell.type != CellValueType.IS_STRING :
                has_header = False
                break
            sen_cell = cells.get(begin_row_index+1,column_index)
            if cell.type != sen_cell.type :
                break
        pass
    
    def __dataframe_import_cells(self,dataframe: pd.DataFrame , cells: Cells , first_row : int , first_column: int):
        column_count = 0
        row_count = 0
        column_index = first_column        
        for column_name in dataframe.columns:  
            column_count = column_count + 1
            row_index = first_row
            cell = cells.get(row_index , column_index )
            cell.put_value(column_name)
            
            row_index = row_index + 1
            for column_value in dataframe[column_name]:
                cell = cells.get(row_index , column_index )
                cell.put_value(column_value)
                row_index = row_index + 1
            column_index = column_index +1
            row_count = row_index - first_row
        return (first_row,first_column,row_count, column_count)
        pass
    
    def __get_dataframe(self , cells : Cells , begin_row_index : int , begin_column_index : int , end_row_index : int , end_column_index : int , has_header: bool, has_total : bool):
        column_title_list =[]
        row_index = 0
        cells_helper = CellsHelper
        if has_header :
            row_index = begin_row_index
        for column_index in range(begin_column_index , end_column_index + 1 ):
            if has_header :
                column_title_list.append (cells.get(row_index,column_index).display_string_value )
            else:
                column_title_list.append (cells_helper.column_index_to_name(column_index) )                       

        start_row = 0
        end_row = 0              
        if has_header :
            start_row = begin_row_index + 1
        else:
            start_row = begin_row_index
        
        if has_total:
            end_row = end_row_index 
        else:
            end_row = end_row_index + 1
                 
        position = 0
        data = {}
        for column_index in range(begin_column_index , end_column_index + 1 ):
            column_data = []
            for row_index in range(start_row ,end_row ):
                column_data.append(cells.get(row_index,column_index).value)
            data[column_title_list[position]] = column_data
            position = position + 1
        return pd.DataFrame(data)
        pass

    def __parse_data_source( self , value : str):        
        matchObj = re.match( r'^=(.*)!\$(.*)\$(\d+):\$(.*)\$(\d+)', value, re.M|re.I)
        if matchObj == None :
            return None
        
        return (matchObj.group(1) , matchObj.group(3) - 1 , CellsHelper.column_name_to_index (matchObj.group(2)),  matchObj.group(5) -1 ,  CellsHelper.column_name_to_index (matchObj.group(4)) )
    def __get_type(self , chart_type : ChartType) -> str : 
        match  chart_type:
            case ChartType.Area:
                return "area"
            case ChartType.AreaStacked:
                return "area"
            case ChartType.Area100PercentStacked:
                return "area"
            case ChartType.Area3D:
                return "area"
            case ChartType.Area3DStacked:
                return "area"
            case ChartType.Area3D100PercentStacked:
                return "area"
            case ChartType.Bar:
                return "bar"
            case ChartType.BarStacked:
                return "bar"
            case ChartType.Bar100PercentStacked:
                return "bar"
            case ChartType.Bar3DClustered:
                return "bar"
            case ChartType.Bar3DStacked:
                return "bar"
            case ChartType.Bar3D100PercentStacked:
                return "bar"
            case ChartType.Bubble:
                return "scatter"
            case ChartType.Bubble3D:
                return "scatter"
            case ChartType.Column:
                return "hist"
            case ChartType.ColumnStacked:
                return "hist"
            case ChartType.Column100PercentStacked:
                return "hist"
            case ChartType.Column3D:
                return "hist"
            case ChartType.Column3DClustered:
                return "hist"
            case ChartType.Column3DStacked:
                return "hist"
            case ChartType.Column3D100PercentStacked:
                return "hist"
            case ChartType.Cone:
                return "scatter"
            case ChartType.ConeStacked:
                return "scatter"
            case ChartType.Cone100PercentStacked:
                return "scatter"
            case ChartType.ConicalBar:
                return "bar"
            case ChartType.ConicalBarStacked:
                return "bar"
            case ChartType.ConicalBar100PercentStacked:
                return "bar"
            case ChartType.ConicalColumn3D:
                return "bar"
            case ChartType.Cylinder:
                return "bar"
            case ChartType.CylinderStacked:
                return "bar"
            case ChartType.Cylinder100PercentStacked:
                return "bar"
            case ChartType.CylindricalBar:
                return "bar"
            case ChartType.CylindricalBarStacked:
                return "bar"
            case ChartType.CylindricalBar100PercentStacked:
                return "bar"
            case ChartType.CylindricalColumn3D:
                return "bar"
            case ChartType.Doughnut:
                return "pie"
            case ChartType.DoughnutExploded:
                return "pie"
            case ChartType.Line:
                return "plot"
            case ChartType.LineStacked:
                return "plot"
            case ChartType.Line100PercentStacked:
                return "plot"
            case ChartType.LineWithDataMarkers:
                return "plot"
            case ChartType.LineStackedWithDataMarkers:
                return "plot"
            case ChartType.Line100PercentStackedWithDataMarkers:
                return "plot"
            case ChartType.Line3D:
                return "plot"
            case ChartType.Pie:
                return "pie"
            case ChartType.Pie3D:
                return "pie"
            case ChartType.PiePie:
                return "pie"
            case ChartType.PieExploded:
                return "pie"
            case ChartType.Pie3DExploded:
                return "pie"
            case ChartType.PieBar:
                return "pie"
            case ChartType.Pyramid:
                return "hist"
            case ChartType.PyramidStacked:
                return "hist"
            case ChartType.Pyramid100PercentStacked:
                return "hist"
            case ChartType.PyramidBar:
                return "hist"
            case ChartType.PyramidBarStacked:
                return "hist"
            case ChartType.PyramidBar100PercentStacked:
                return "hist"
            case ChartType.PyramidColumn3D:
                return "hist"
            case ChartType.Radar:
                return "hist"
            case ChartType.RadarWithDataMarkers:
                return "hist"
            case ChartType.RadarFilled:
                return "hist"
            case ChartType.Scatter:
                return "scatter"
            case ChartType.ScatterConnectedByCurvesWithDataMarker:
                return "scatter"
            case ChartType.ScatterConnectedByCurvesWithoutDataMarker:
                return "scatter"
            case ChartType.ScatterConnectedByLinesWithDataMarker:
                return "scatter"
            case ChartType.ScatterConnectedByLinesWithoutDataMarker:
                return "scatter"
            case ChartType.StockHighLowClose:
                return "boxplot"
            case ChartType.StockOpenHighLowClose:
                return "boxplot"
            case ChartType.StockVolumeHighLowClose:
                return "boxplot"
            case ChartType.StockVolumeOpenHighLowClose:
                return "boxplot"
            case ChartType.Surface3D:
                return "plot_surface"
            case ChartType.SurfaceWireframe3D:
                return "plot_surface"
            case ChartType.SurfaceContour:
                return "plot_surface"
            case ChartType.SurfaceContourWireframe:
                return "plot_surface"
            case ChartType.BoxWhisker:
                return "hist"
            case ChartType.Funnel:
                return "hist"
            case ChartType.ParetoLine:
                return "hist"
            case ChartType.Sunburst:
                return "hist"
            case ChartType.Treemap:
                return "hist"
            case ChartType.Waterfall:
                return "hist"
            case ChartType.Histogram:
                return "hist"
            case ChartType.Map:
                return "hist"
            case ChartType.RadialHistogram:
                return "hist"