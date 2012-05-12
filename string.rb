class String
  
  def sanitize
    self.gsub(/\\/, '/').gsub(/#/, '\#').gsub(/'/, %Q|\'|).gsub(/"/, %q|\"|)
  end

  def to_variable
  
    self.downcase.gsub(/%/, '').strip.gsub(/-/, '_').gsub(/ /, '_').gsub(/__/, '_')
  end
  
  def to_excel_number
    a = self.unpack("C*").map{|c| c - 65}
    return a[0].to_i if a.size == 1
    raise "too many excel sheet column number" if a.size > 2
    
    (a[0] + 1) * 25 + a[1]
  end
end