#-----------------------------------------------------------------------------------------------------
# console help
#-----------------------------------------------------------------------------------------------------
# include help text to code for future compilation
HELP=
    %{
Опции
-----
  --help или -h      справочная информация

Параметры интеграции
--------------------
Перед началом работы нужно открыть файл MS Project, который будем синхронизировать.
Параметры синхронизации задаются в задаче с названием 'Синхронизация Redmine'.

Файл в формате yml в кодировке windows-1251, состоит из строк вида
  имяпараметра: значениепараметра
Имя параметра - без черточек. Если значение параметра содержит пробелы, писать в кавычках, например:
  base: karat_2013
  user: "Пупкин ВВ"
Параметры командной строки имеют приоритет перед параметрами файла настроек.

}

#-----------------------------------------------------------------------------------------------------
# exit on condition
#-----------------------------------------------------------------------------------------------------
def chk cond,msg;if cond;puts msg+"\n\n";exit 0;end;end

#-----------------------------------------------------------------------------------------------------
# REST to Redmine API
#-----------------------------------------------------------------------------------------------------
def rm_request path, data = nil,method = nil, verbose = false
  reply = nil
  hdrs = {'Content-Type'=>'application/json', 'X-Redmine-API-Key' => $settings['redmine_api_key']}
  reply = Net::HTTP.start($settings['redmine_host'], $settings['redmine_port']) do |http|
    if data
      if !method || method == 'POST'
        puts "POST #{path}" if verbose
        http.request_post(path,data.to_json, hdrs)
      elsif method == 'PUT'
        # code for PUT here
        puts "PUT #{path}" if verbose
        http.send_request 'PUT', path, JSON.unparse(data), hdrs
      end
    else
      puts "GET #{path}" if verbose
      http.request_get(path, hdrs)
    end
  end
  return reply
end

#-----------------------------------------------------------------------------------------------------
# REST get from Redmine API
#-----------------------------------------------------------------------------------------------------
def rm_get path, entity, msg

  rmp = nil
  re = rm_request path
  chk (re.code!='200'), msg + "\n#{re.code} #{re.msg}\n\n"
  rmp = JSON.parse(re.body)[entity] rescue nil
  chk !rmp, msg + "\n#{re.body.inspect}\n\n"
  return rmp

end

#-----------------------------------------------------------------------------------------------------
# REST create to Redmine API
#-----------------------------------------------------------------------------------------------------
def rm_create path, entity, data, msg

  rmp = nil
  re = rm_request path, entity => data
  chk (re.code!='201'), msg + "\n#{re.code} #{re.msg}\n\n"
  rmp = JSON.parse(re.body)[entity] rescue nil
  chk !rmp, msg + "\n#{re.body.inspect}\n\n"
  return rmp

end

#-----------------------------------------------------------------------------------------------------
# REST update to Redmine API
#-----------------------------------------------------------------------------------------------------
def rm_update path, data, msg

  re = rm_request path, data, 'PUT'
  puts re.body
  chk (re.code!='200'), msg + "\n#{re.code} #{re.msg}\n\n"
  return true

end


