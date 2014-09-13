#!/usr/bin/env ruby
# -*- coding: utf-8 -*-

require 'serialport'

STDOUT.sync = true

def xbee_api_parse_data_sample(buffer)
  digital_sample = {}
  analog_sample = {}
  digital_index = []
  analog_index = []

  digital_channel_mask = buffer[0..1].pack("c*").unpack("n")[0]
  analog_channel_mask = buffer[2]
  buffer.slice!(0, 3)

  (0..12).each{|i|
    digital_index.push(i) if ((digital_channel_mask & 0x1) == 1)
    digital_channel_mask >>= 1
  }
  (0..3).each{|i|
    analog_index.push(i) if ((analog_channel_mask & 0x1) == 1)
    analog_channel_mask >>= 1
  }

  if (digital_index.size != 0) then
    digital_value = buffer[0..1].pack("c*").unpack("n")[0]
    digital_index.each{|i|
      digital_sample[i.to_s.to_i] = (digital_value >> i) & 0x1
    }
    buffer.slice!(0, 2)
  end

  analog_index.each{|i|
    analog_sample[i.to_s.to_i] = buffer[0..1].pack("c*").unpack("n")[0]
    buffer.slice!(0, 2)
  }

  return {
    digital: digital_sample,
    analog: analog_sample,
  }
end

def xbee_api_parse_receive_packet(packet)
  frame = {}

  frame[:frame_type] = packet[0].to_s(16).upcase
  frame[:source_address] = packet[1..8].pack("C*").unpack("H*")[0].upcase
  frame[:network_address] = packet[9..10].pack("C*").unpack("H*")[0].upcase
  frame[:receive_options] = packet[11]
  frame[:number_of_samples] = packet[12]

  return frame
end

def xbee_api_parse_io_data_sample_rx_indicator(packet)
  frame = xbee_api_parse_receive_packet(packet)
  frame[:sample_data] = xbee_api_parse_data_sample(packet[13..(packet.size-1)])
  return frame
end

class XBee
  def initialize(dev)
    @sp = SerialPort.new(dev, 9600, 8, 1,SerialPort::NONE)
    @sp.flow_control = SerialPort::NONE
    @sp.read_timeout = 1000
    @buffer = []
  end

  def unescape(payload)
    ret = []

    escaped = false
    payload.each_with_index{|val, i|
      if (escaped) then
        ret.push(val ^ 0x20)
        escaped = false
      elsif (val == 0x7D) then
        escaped = true
      else
        ret.push(val)
      end
    }
    return ret
  end

  def check_checksum(buffer)
    sum = 0;
    buffer[0..(buffer.size-2)].each{|val|
      sum += val
    }
    sum = 0xFF - (sum & 0xFF)
    return buffer[-1] == sum
  end

  def receive
    while (true)
      data = @sp.read(30)
      if (data != nil)
        @buffer += data.unpack("C*")
      end

      start_i = @buffer.index(0x7E)
      next if start_i == nil

      @buffer.slice!(0, start_i)
      next if @buffer.size < 20

      length = unescape(@buffer)[1..2].pack("c*").unpack("n")[0]
      next if @buffer.size < (3 + length + @buffer.count(0x7D))
      @buffer.slice!(0, 3)

      start_i = @buffer.index(0x7E)
      start_i = @buffer.size if start_i == nil

      packet = unescape(@buffer.slice!(0, start_i))
      
      if check_checksum(packet)
        yield packet
      end
      break
    end
  end
end

xbee = XBee.new('/dev/ttyO4')

def calc_temp(ad_value)
  return (((ad_value * 1200) / 1023) - 424) / 6.25
end

##################################################
data = {}
while (true)
  xbee.receive{|packet|
    begin
      frame = xbee_api_parse_io_data_sample_rx_indicator(packet)

      printf("%s: temperature: %.2f℃ (battery: %s)\n",
             frame[:source_address],
             calc_temp(frame[:sample_data][:analog][3]),
             (frame[:sample_data][:digital][5] == 1) ? 'OK' : 'NG')

    rescue Exception => e
      p e
    end
  }
end
