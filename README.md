# php library for export excel

## usage

```
$data = [
    [
        'id' => 1,
        'name' => 'Jon',
        'age' => 18,
        'header' => './upload/1.jpg',
    ],
    [
        'id' => 2,
        'name' => 'Lucy',
        'age' => 16,
        'header' => './upload/2.jpg',
    ],
];
$headers = ['ID','姓名','年龄','头像'];
$et = new Bingher\ExportTable\ExportTable;
$et->data($data,$headers)->output();
```
