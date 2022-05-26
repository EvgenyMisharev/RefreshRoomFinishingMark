using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RefreshRoomFinishingMark
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class RefreshRoomFinishingMarkCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            Room roomFromParam = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfClass(typeof(SpatialElement))
                .Where(e => e.GetType() == typeof(Room))
                .Cast<Room>()
                .FirstOrDefault(r => r.Area > 0);

            if (roomFromParam == null)
            {
                TaskDialog.Show("Revit", "В проекте нет размещенных помещений!");
                return Result.Cancelled;
            }

            //АР_ТипПола_Плагин
            Guid floorTypePluginParamGUID = new Guid("34eac6db-112c-49fc-a3dc-7e3470b9656f");
            Parameter floorTypePluginParam = roomFromParam.get_Parameter(floorTypePluginParamGUID);
            if(floorTypePluginParam == null)
            {
                TaskDialog.Show("Revit", "В проекте отсутствует параметр \"АР_ТипПола_Плагин\"!");
                return Result.Cancelled;
            }
            //АР_ОтделкаСтен_Плагин
            Guid wallFinishPluginParamGUID = new Guid("326001d8-4e61-494a-8eeb-74e190e11bcf");
            Parameter wallFinishPluginParam = roomFromParam.get_Parameter(wallFinishPluginParamGUID);
            if (wallFinishPluginParam == null)
            {
                TaskDialog.Show("Revit", "В проекте отсутствует параметр \"АР_ОтделкаСтен_Плагин\"!");
                return Result.Cancelled;
            }
            //АР_ОтделкаСтенСнизу_Плагин
            Guid bottomWallFinishPluginParamGUID = new Guid("ee4a5627-a73d-40fa-80a9-d8702abb2f89");
            Parameter bottomWallFinishPluginParam = roomFromParam.get_Parameter(bottomWallFinishPluginParamGUID);
            if (bottomWallFinishPluginParam == null)
            {
                TaskDialog.Show("Revit", "В проекте отсутствует параметр \"АР_ОтделкаСтенСнизу_Плагин\"!");
                return Result.Cancelled;
            }
            //АР_ОтделкаПотолка_Плагин
            Guid ceilingFinishPluginParamGUID = new Guid("08ea6382-373d-4347-b723-4eab2427d250");
            Parameter ceilingFinishPluginParam = roomFromParam.get_Parameter(ceilingFinishPluginParamGUID);
            if (ceilingFinishPluginParam == null)
            {
                TaskDialog.Show("Revit", "В проекте отсутствует параметр \"АР_ОтделкаПотолка_Плагин\"!");
                return Result.Cancelled;
            }

            //АР_ТипПола_Ключ
            Parameter floorTypeKeyParam = roomFromParam.LookupParameter("АР_ТипПола_Ключ");
            if (floorTypeKeyParam == null)
            {
                TaskDialog.Show("Revit", "В проекте отсутствует параметр \"АР_ТипПола_Ключ\"!");
                return Result.Cancelled;
            }
            //АР_ОтделкаСтен_Ключ
            Parameter wallFinishKeyParam = roomFromParam.LookupParameter("АР_ОтделкаСтен_Ключ");
            if (wallFinishKeyParam == null)
            {
                TaskDialog.Show("Revit", "В проекте отсутствует параметр \"АР_ОтделкаСтен_Ключ\"!");
                return Result.Cancelled;
            }
            //АР_ОтделкаСтенСнизу_Ключ
            Parameter bottomWallFinishKeyParam = roomFromParam.LookupParameter("АР_ОтделкаСтенСнизу_Ключ");
            if (bottomWallFinishKeyParam == null)
            {
                TaskDialog.Show("Revit", "В проекте отсутствует параметр \"АР_ОтделкаСтенСнизу_Ключ\"!");
                return Result.Cancelled;
            }
            //АР_ОтделкаПотолка_Ключ
            Parameter ceilingFinishKeyParam = roomFromParam.LookupParameter("АР_ОтделкаПотолка_Ключ");
            if (ceilingFinishKeyParam == null)
            {
                TaskDialog.Show("Revit", "В проекте отсутствует параметр \"АР_ОтделкаПотолка_Ключ\"!");
                return Result.Cancelled;
            }

            List<Room> roomList = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .OfClass(typeof(SpatialElement))
                .Where(e => e.GetType() == typeof(Room))
                .Cast<Room>()
                .Where(r => r.Area > 0)
                .ToList();
            using(Transaction t = new Transaction(doc))
            {
                t.Start("Обновление марок помещений");
                foreach (Room room in roomList)
                {
                    if (room.LookupParameter("АР_ТипПола_Ключ").AsElementId().IntegerValue == -1)
                    {
                        room.get_Parameter(floorTypePluginParamGUID).Set("");
                    }
                    else
                    {
                        room.get_Parameter(floorTypePluginParamGUID).Set(room.LookupParameter("АР_ТипПола_Ключ").AsValueString());
                    }

                    if(room.LookupParameter("АР_ОтделкаСтен_Ключ").AsElementId().IntegerValue == -1)
                    {
                        room.get_Parameter(wallFinishPluginParamGUID).Set("");
                    }
                    else
                    {
                        room.get_Parameter(wallFinishPluginParamGUID).Set(room.LookupParameter("АР_ОтделкаСтен_Ключ").AsValueString());
                    }

                    if(room.LookupParameter("АР_ОтделкаСтенСнизу_Ключ").AsElementId().IntegerValue == -1)
                    {
                        room.get_Parameter(bottomWallFinishPluginParamGUID).Set("");
                    }
                    else
                    {
                        room.get_Parameter(bottomWallFinishPluginParamGUID).Set(room.LookupParameter("АР_ОтделкаСтенСнизу_Ключ").AsValueString());
                    }
                    
                    if(room.LookupParameter("АР_ОтделкаПотолка_Ключ").AsElementId().IntegerValue == -1)
                    {
                        room.get_Parameter(ceilingFinishPluginParamGUID).Set("");
                    }
                    else
                    {
                        room.get_Parameter(ceilingFinishPluginParamGUID).Set(room.LookupParameter("АР_ОтделкаПотолка_Ключ").AsValueString());
                    }
                }
                t.Commit();
            }

            TaskDialog.Show("Revit","Обработка завершена!");
            return Result.Succeeded;
        }
    }
}
