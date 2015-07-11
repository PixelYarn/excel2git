#!/usr/bin/env ruby

require 'roo'
require 'csv'
require 'git'
require 'pathname'
require 'thor'
require 'pry'

class ExcelToGit < Thor
  desc "init", "create a new git repo in the current directory"
  def init
    Git.init
  end

  desc "save SOURCE.xlsx DESTINATION.csv", "Export source to destination and commit to git"
  def save(source, destination)
    puts 'executing save...'
    excel = Roo::Excelx.new(source)
    outfile = Pathname.new(destination)
    output = File.open(outfile.to_path, "w")

    puts 'exporting source'
    1.upto(excel.last_row) do |line|
      output.write(CSV.generate_line(excel.row(line)))
    end

    puts 'connecting to repo'
    repo = Git.open(outfile.dirname)
    resp = repo.add(outfile.to_path)
    puts resp unless resp.strip.empty?
    begin
      resp = repo.commit("Update #{outfile.basename}")
      puts resp unless resp.strip.empty?
    rescue
    end
  end
end

ExcelToGit.start(ARGV)
