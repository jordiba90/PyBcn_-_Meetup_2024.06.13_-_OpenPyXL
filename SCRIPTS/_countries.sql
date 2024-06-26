SELECT TOP (1000) [OrderID]
      ,[CustomerID]
      ,[EmployeeID]
      ,FORMAT([OrderDate],'dd/MM/yyyy','es-ES') as [OrderDate]
      ,FORMAT([ShippedDate],'dd/MM/yyyy','es-ES') as [ShippedDate]
      ,FORMAT([RequiredDate],'dd/MM/yyyy','es-ES') as [RequiredDate]
      ,[ShipVia]
      ,[Freight]
      ,[ShipName]
      ,[ShipAddress]
      ,[ShipCity]
      ,[ShipRegion]
      ,[ShipPostalCode]
      ,[ShipCountry]
  FROM [Northwind].[dbo].[Orders]