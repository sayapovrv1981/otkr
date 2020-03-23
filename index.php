<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
	<head>
		<meta http-equiv="content-type" content="text/html; charset=utf-8" />
		<meta name="description" content="" />
		<meta name="keywords" content="" />
		<title>Загрузка xlsx</title>
		<script type="text/javascript" src=" https://code.jquery.com/jquery-1.11.2.js "></script>
		<link rel="stylesheet"
					href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css"
					integrity="sha384-BVYiiSIFeK1dGmJRAkycuHAHRg32OmUcww7on3RYdg4Va+PmSTsz/K68vbdEjh4u" crossorigin="anonymous">
<?php
require_once("config.php");
ini_set('upload_max_filesize', MAX_FILE_SIZE); //ограничение в 3 мб  устанавливается в config

?>
</head>
  <body>
  <h1>Форма для загрузки файла .xlsx и его обработки</h1>
  <form method="post" action="" enctype="multipart/form-data">
  <label for="inputfile">Загрузка файла</label>
  <input type="file" id="inputfile" name="inputfile"></br>
  <input type="submit" value="Загрузить">
  </form>

<?php
if (isset($_POST))
{
  //print_r($_POST);
  //echo "files:<br>";
  //print_r($_FILES);
  //echo "const UPLOAD_ERR_OK:",UPLOAD_ERR_OK,"MAX_FILE_SIZE",MAX_FILE_SIZE;
  if ($_FILES['inputfile']['error'] == UPLOAD_ERR_OK) //проверка на наличие ошибок при загрузке файла

  {
    if (stripos($_FILES['inputfile']['name'],'.xlsx')>0) // проверяем формат xlsx
    {
    //  echo $_FILES['inputfile']['name'];
      $destiation_dir =  __DIR__  .  DIRECTORY_SEPARATOR   . $_FILES['inputfile']['name']; // директория для размещения файла
      if (is_writable (__DIR__  .  DIRECTORY_SEPARATOR )) // проверка текущего директория на права доступа
      {
        if (move_uploaded_file($_FILES['inputfile']['tmp_name'], $destiation_dir)) //перемещение в указанную директорию
        {
          //echo 'Файл успешно загружен!'; //оповещаем пользователя об успешной загрузке файла
          $excel = PHPExcel_IOFactory::load($destiation_dir); // создаем объект файла xlsx

          if ($excel->getSheetCount()>1)  //проверяем число листов в xlsx файле
          {
            if ($excel->sheetNameExists('first'))  //проверяем наличие листа first xlsx файле
            {
              $excel->setActiveSheetIndexByName('first');
              if ((! empty($excel->getActiveSheet()->getCell('A1')->getValue()))&&(! empty($excel->getActiveSheet()->getCell('B1')->getValue()))&&(! empty($excel->getActiveSheet()->getCell('C1')->getValue())))
              {
                //  загрузка данных листа first  в таблицу clients (каждый раз пересоздается в БД);
                  echo excel2mysql($excel->getActiveSheet(), $db_link, clients,0,["id","client_name","balance"]) ? "<br>Загрузка данных листа first успешно завершена\n" : "<br>Ошибка при загрузке данных листа first\n";
              }
              else
              {
                echo "Не выполнено условие заполнения столбцов 1-3  в листе first";
              }
            }
            else
            {
              echo "Ошибка лист с именем 'first' отсутствует в загружаемом файле";
              exit;
            }

            if ($excel->sheetNameExists('second')) //проверяем наличие листа second xlsx файле
            {
              $excel->setActiveSheetIndexByName('second');
              if ((! empty($excel->getActiveSheet()->getCell('A1')->getValue()))&&(! empty($excel->getActiveSheet()->getCell('B1')->getValue())))
              {
                //  загрузка данных листа second  в таблицу cash (каждый раз пересоздается в БД);
                echo excel2mysql($excel->getActiveSheet(), $db_link, cash,0,["id","cash_deals"]) ? "<br>Загрузка данных листа second успешно завершена\n" : "<br>Ошибка при загрузке данных листа second\n";
              }
              else
              {
                echo "Не выполнено условие заполнения столбцов 1-2 в листе second";
              }

            }
            else
            {
              echo "Ошибка лист с именем 'second' отсутствует в загружаемом файле";
              exit;
            }
            if ($result = mysqli_query($db_link,"SELECT clients.id, clients.client_name, clients.balance+casht.sum as cash_balance FROM clients, (SELECT id, SUM(cash_deals) as sum FROM cash GROUP BY id) as casht WHERE clients.id = casht.id"));
              {//формируется выборка
                $result = mysqli_fetch_all($result, MYSQLI_ASSOC);

                  ?>
                  <table class="table">
                    <thead>
                      <tr>
                        <th scope="col">
                            <p>#</p>
                        </th>
                        <th scope="col">
                            <p>ФИО клиента</p>
                        </th>
                        <th>
                            <p>Текущий остаток с учетом вводов/выводов</p>
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                    <?php
                    if (count($result)>0)
                    {
                    	foreach($result as $row)
                    	{
                    		echo '<tr>
                    						<td>'.$row['id'].'</td>
                    						<td>'.$row['client_name'].'</td>
                    						<td>'.$row['cash_balance'].'</td>
                    					</tr>';
                    	}
                    }
                    ?>
                    </tbody>
                  </table>
                </body>
              </html>
                  <?php
              }
          }
          else
          {
            echo 'Загрузка файла не возможна, в нем количество листов не равно двум';
          }
        }
        else
        {
          echo 'Загрузка файла не удалась';
        }
      }
      else {
        echo 'Отсутствуют права доступа к папке/файлу: '.__DIR__  .  DIRECTORY_SEPARATOR;
      }

    }
    else {
      {
        echo 'Формат файла не поддерживается';
      }
    }

  }
  else
  {
    switch ($_FILES['inputfile']['error'])
    {
      case UPLOAD_ERR_FORM_SIZE:
        case UPLOAD_ERR_INI_SIZE:
          echo 'Превышен размер загружаемого файла';
          brake;
        case UPLOAD_ERR_NO_FILE:
          echo 'Отсутствует файл для загрузки';
        break;
      default:
        echo 'Что-то пошло не так';
    }
  }
}
