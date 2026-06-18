from django.contrib import admin
from django.urls import path,include
from GSTapp import views

app_name = 'view'

urlpatterns = [
    path('',views.index,name='index'),
    path('tallyxml',views.tallyxml,name='tally_xml'),
    # path('master_ledger',views.master_ledger,name='master_ledger'),
    # path('master_duties',views.master_duties,name='master_duties'),
    path('master',views.master,name='master'),
    path('pur_sale',views.Purchase_Sales.as_view(),name='pur_sale'),
    path('pay_con_rec',views.Pay_Con_Rec.as_view(),name='pay_con_rec'),
    # path('master',views.Master.as_view(),name='master'),
    path('master_ledger',views.Master_Ledger.as_view(),name='master_ledger'),
    path('master_duties',views.Master_Duties.as_view(),name='master_duties'),
    path('master_ps',views.Master_PS.as_view(),name='master_ps'),
    path('jsonexcel',views.jsonexcel,name='json_excel'),
    path('download-excel/', views.download_excel, name='download_excel'),
    path('download_excel_pay_con_rec/', views.download_excel_pay_con_rec, name='download_excel_pay_con_rec'),
    path('download_excel_master_1/', views.download_excel_master_1, name='download_excel_master_1'),
    path('download_excel_master_2/', views.download_excel_master_2, name='download_excel_master_2'),
    path('download_excel_master_3/', views.download_excel_master_3, name='download_excel_master_3'),
    path('download_xml/', views.download_xml, name='download_xml'),
    path('download_xml_file/', views.download_xml_file, name='download_xml_file'),
    path('download_xml_excel/', views.download_xml_excel, name='download_xml_excel')
]