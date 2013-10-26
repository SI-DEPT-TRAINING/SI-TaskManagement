proj_root_dir = File.expand_path("../../../../", __FILE__)
worker_processes 4
listen      "#{proj_root_dir}/SI-TaskManagement/tmp/sockets/universe.sock"
pid         "#{proj_root_dir}/SI-TaskManagement/tmp/pids/unicorn.pid"
stderr_path "#{proj_root_dir}/SI-TaskManagement/log/unicorn.log"
stdout_path "#{proj_root_dir}/SI-TaskManagement/log/unicorn.log"
