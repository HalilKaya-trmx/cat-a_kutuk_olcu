import win32com.client

# CATIA başlatma
CATIA = win32com.client.Dispatch("CATIA.Application")
active_doc = CATIA.ActiveDocument
part = active_doc.Part
hybrid_shape_factory = part.HybridShapeFactory
hybrid_bodies = part.HybridBodies
new_hybrid_body = hybrid_bodies.Add()
new_hybrid_body.Name = "Rough stock"

# Parçadan uzakta olan 6 başlangıç düzlemi oluşturma
planes_equations = [
    (-1.0, 0.0, 0.0, 33333.0),
    (1.0, 0.0, 0.0, 33333.0),
    (0.0, -1.0, 0.0, 33333.0),
    (0.0, 1.0, 0.0, 33333.0),
    (0.0, 0.0, -1.0, 33333.0),
    (0.0, 0.0, 1.0, 33333.0)
]
initial_planes = [hybrid_shape_factory.AddNewPlaneEquation(*eq) for eq in planes_equations]
for plane in initial_planes:
    new_hybrid_body.AppendHybridShape(plane)
    plane.Compute()

# Başlangıç düzlemleri ve seçilen parça arasındaki mesafeleri hesaplama
spa_workbench = active_doc.GetWorkbench("SPAWorkbench")
selection = active_doc.Selection
status = selection.SelectElement2(["Body", "Body"], "Select body to be measured, <Esc> ... exit.", False)
selected_element = selection.Item(1).Value
reference_body = part.CreateReferenceFromObject(selected_element)

distances = []
for plane in initial_planes:
    measurable = spa_workbench.GetMeasurable(plane)
    distance = measurable.GetMinimumDistance(reference_body)
    distances.append(distance)

# Teğet düzlemleri oluşturma
def create_single_tangent_plane(reference, distance, normal):
    if normal in [(1.0, 0.0, 0.0), (0.0, 1.0, 0.0), (0.0, 0.0, 1.0)]:
        plane = hybrid_shape_factory.AddNewPlaneOffset(reference, -distance, False)
    else:
        plane = hybrid_shape_factory.AddNewPlaneOffset(reference, -distance, False)
    return plane

tangent_planes = []
for plane, dist, eq in zip(initial_planes, distances, planes_equations):
    tangent_planes.append(create_single_tangent_plane(plane, dist, eq[:3]))

for tangent_plane in tangent_planes:
    new_hybrid_body.AppendHybridShape(tangent_plane)
    tangent_plane.Compute()
    print(tangent_plane.Name)
    
    
    
# Kütük ölçülerini belirleme
# KutukGenislik = abs(distances[1] - distances[0])
# KutukYukseklik = abs(distances[3] - distances[2])
# KutukDerinlik = abs(distances[5] - distances[4])
# KutukOlculeri = f"{round(KutukGenislik, 2)}x{round(KutukYukseklik, 2)}x{round(KutukDerinlik, 2)}"
# new_hybrid_body.Name = f"Rough stock {KutukOlculeri}"

# # Teğet düzlemler arasındaki mesafeyi hesaplama
# for i in range(len(tangent_planes) - 1):
#     for j in range(i + 1, len(tangent_planes)):
#         plane1 = tangent_planes[i]
#         plane2 = tangent_planes[j]
#         measurable1 = spa_workbench.GetMeasurable(plane1)
#         distance = measurable1.GetMinimumDistance(plane2)
#         if distance > 0:  # sıfırdan büyük mesafeleri kontrol edin
#             print(f"Mesafe {plane1.Name} ile {plane2.Name} arasında: {distance}")

distance_list = []

for i in range(len(tangent_planes) - 1):
    for j in range(i + 1, len(tangent_planes)):
        plane1 = tangent_planes[i]
        plane2 = tangent_planes[j]
        measurable1 = spa_workbench.GetMeasurable(plane1)
        distance = measurable1.GetMinimumDistance(plane2)
        if distance > 0:  # sıfırdan büyük mesafeleri kontrol edin
            distance_list.append(distance)

# Mesafeleri sıralama ve ilk 5 değeri alma
sorted_distances = sorted(distance_list)
top_5_distances = sorted_distances[:5]

# Değerleri birleştirme
dimension_value = 'x'.join([str(round(dist, 2)) for dist in top_5_distances])

parameters = part.Parameters
if "Dimension" not in [param.Name for param in parameters]:
    strParam = parameters.CreateString("Dimension", dimension_value)
    # part.HybridBodies.DeleteObject("Rought stock")
else:
    strParam = parameters.Item("Dimension")
    strParam.Value = dimension_value
    # part.HybridBodies.DeleteObject("Rought stock")
    

rough_stock_body = part.HybridBodies.Item("Rough stock")
part.HybridBodies.DeleteObject(rough_stock_body)



# parameters = part.Parameters
# if "Dimension" not in [param.Name for param in parameters]:
#     strParam = parameters.CreateString("Dimension", KutukOlculeri)
# else:
#     strParam = parameters.Item("Dimension")
#     strParam.Value = KutukOlculeri
