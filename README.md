# TBM Position Check
TBM Position Check สำหรับการเซ็ตตำแหน่งของ Tunnel Boring Mechine (TBM) ก่อนเริ่มเจาะ ซึ่งวิศวกรสำรวจควรจะต้องคำนวณตรวจสอบตำแหน่ง TBM เทียบกับระบบนำทาง TBM Guidance System เพื่อยืนยันตำแหน่ง

## TBM Guidance System Overview
![01](https://github.com/user-attachments/assets/e8b481ff-a198-4cf6-8e9c-733e28d2e188)
 _VBA Excel สำหรับคำนวณตรวจสอบ_

![02](https://github.com/user-attachments/assets/e15e82c6-da8b-4f3d-ae4d-53a20c6bbeec)
_Enzan Guidance System ระบบนำทางที่อยู่ใน TBM_

## Workflow
_**Video Preview :**_ [Youtube](https://www.youtube.com/watch?v=RnKc08XiZW0)\
_**VBA Script :**_ [ManualTBMPosition_R4.bas](https://github.com/suben-mk/TBM-Position-Check-EzanGuidance-System/blob/main/ManualTBMPosition_R4.bas)\
_**VBA Excel Program :**_ [Manual TBM Position Program Rev.04](https://github.com/suben-mk/TBM-Position-Check-EzanGuidance-System/blob/main/Manual%20TBM%20Position%20Program%20Rev.04.xlsm)

 1. เก็บข้อมูลภาคสนาม\
    1.1 As-built TBM body (Wriggle Survey) 4 Sections
    
    ![03-1](https://github.com/user-attachments/assets/ddcb8166-d43a-414b-a5bb-d92669ca7656)

    ![04-5](https://github.com/user-attachments/assets/c31f3707-a0a7-4b5b-ad91-a85a9a143e3e)

    1.2 TBM Rolling 
    
    ![05-1](https://github.com/user-attachments/assets/e9232bd7-3497-4182-884b-4c478e9ea88e)

    1.3 TBM Targets

    ![06](https://github.com/user-attachments/assets/32a103f3-1d18-4469-89a9-0a7286867e27)

    1.4 Articulation Jack Stroke

    ![07](https://github.com/user-attachments/assets/8933d338-0974-4a3a-87a9-be0318fb0b39)

 2. ขั้นตอนคำนวณตำแหน่ง TBM\
    2.1 TBM center axis (TBM rear axis และ TBM front axis) หาพิกัด center แต่ละ Section จะใช้วิธี Line of Best Fit (E,N) และ Best-fit circle (Y,Z) ของ the Kasa method

    ![08](https://github.com/user-attachments/assets/c58791f8-19a4-4b06-9480-7df186885c73)

    2.2 TBM azimuth, TBM pitching และ TBM rolling

    ![12](https://github.com/user-attachments/assets/8010fab9-acb0-4e0e-a067-3ad68074b80a)

    2.3 TBM targets ซึ่งจะแปลงระบบพิกัดจากโครงการ (N,E,Z) เป็นระบบพิกัด TBM local coordinates (MX,MY,MZ) ที่ TBM azimuth, pitching, rolling = 0 องศา

    ![09](https://github.com/user-attachments/assets/f577c03d-f9ad-48bd-9eef-158f9e666bb1)

    2.4 TBM parameter ระยะ design ต่างๆจาก TBM Drawing และระยะยืด Articulation jack stroke ที่วัดได้จะเป็นค่าตั้งต้น (Set 0) เทียบกับการวัดครั้งต่อไป ซึ่งจะคำนวณหามุมราบและมุมดิ่งระหว่างแกน TBM rear axis และ TBM front axis

    ![10](https://github.com/user-attachments/assets/95bc72f9-410b-4834-a540-1cec2236f78e)

    2.5 เตรียมข้อมูล Tunnel Alignment ทุกๆ Chainage = 50 cm. (Ch,N,E,Z)

    ![13](https://github.com/user-attachments/assets/7e57d19c-3a60-4344-b706-8e584b3881cf)

    2.6 คำนวณตำแหน่ง TBM (ข้อมูลสีน้ำเงิน) ที่ Tail, Articulation, Head ตามระยะ TBM Drawing และเทียบค่า Deviation จาก Tunnel Alignment

    ![11](https://github.com/user-attachments/assets/6ce12c14-e479-4d35-a4e1-ad01e6fcd628)

    



    
