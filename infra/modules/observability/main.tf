# Log retention and the alarms that say something is wrong.
#
# The log group is created here rather than being left to Lambda, because a
# group Lambda creates on its own keeps logs forever. Owning it is the only way
# to put a retention period on it.

terraform {
  required_version = ">= 1.11"

  required_providers {
    aws = {
      source  = "hashicorp/aws"
      version = "~> 6.0"
    }
  }
}

locals {
  log_group_name = "/aws/lambda/${var.function_name}"

  # a metric filter has to put its counts somewhere, and keeping them out of
  # AWS/Lambda avoids any confusion with the metrics Lambda publishes itself
  metric_namespace = "AccessibilityEnhancer/${var.environment}"
}

resource "aws_cloudwatch_log_group" "lambda" {
  name              = local.log_group_name
  retention_in_days = var.log_retention_days
}

resource "aws_sns_topic" "alarms" {
  name = "${var.function_name}-alarms"
}

# each address gets a confirmation email from AWS and stays "pending" until
# someone clicks the link, which Terraform cannot do for them
resource "aws_sns_topic_subscription" "email" {
  for_each = toset(var.alarm_email_addresses)

  topic_arn = aws_sns_topic.alarms.arn
  protocol  = "email"
  endpoint  = each.value
}

# a deck that fails outright: a corrupt upload, a bucket permission that is
# wrong, or an unhandled crash in the pipeline
resource "aws_cloudwatch_metric_alarm" "errors" {
  alarm_name          = "${var.function_name}-errors"
  alarm_description   = "The pipeline raised an unhandled error on at least one deck."
  namespace           = "AWS/Lambda"
  metric_name         = "Errors"
  statistic           = "Sum"
  period              = 300
  evaluation_periods  = 1
  threshold           = var.error_threshold
  comparison_operator = "GreaterThanOrEqualToThreshold"

  dimensions = {
    FunctionName = var.function_name
  }

  # no invocations at all is the normal state overnight, not a problem
  treat_missing_data = "notBreaching"

  alarm_actions = [aws_sns_topic.alarms.arn]
  ok_actions    = [aws_sns_topic.alarms.arn]
}

# reserved concurrency is doing its job, but someone is waiting on a deck that
# is being turned away
resource "aws_cloudwatch_metric_alarm" "throttles" {
  alarm_name          = "${var.function_name}-throttles"
  alarm_description   = "Uploads are arriving faster than the reserved concurrency allows, so some are being throttled."
  namespace           = "AWS/Lambda"
  metric_name         = "Throttles"
  statistic           = "Sum"
  period              = 300
  evaluation_periods  = 1
  threshold           = 1
  comparison_operator = "GreaterThanOrEqualToThreshold"

  dimensions = {
    FunctionName = var.function_name
  }

  treat_missing_data = "notBreaching"

  alarm_actions = [aws_sns_topic.alarms.arn]
  ok_actions    = [aws_sns_topic.alarms.arn]
}

# a run getting close to the timeout is the warning before decks start being
# cut off partway through, which loses the work rather than failing cleanly
resource "aws_cloudwatch_metric_alarm" "duration" {
  alarm_name          = "${var.function_name}-approaching-timeout"
  alarm_description   = "A deck took most of the function's timeout. The next one slightly larger will be cut off."
  namespace           = "AWS/Lambda"
  metric_name         = "Duration"
  statistic           = "Maximum"
  period              = 300
  evaluation_periods  = 1
  threshold           = var.timeout_seconds * 1000 * var.duration_alarm_fraction
  comparison_operator = "GreaterThanThreshold"

  dimensions = {
    FunctionName = var.function_name
  }

  treat_missing_data = "notBreaching"

  alarm_actions = [aws_sns_topic.alarms.arn]
  ok_actions    = [aws_sns_topic.alarms.arn]
}

# The pipeline catches a captioning failure and writes a placeholder rather than
# failing the deck, so this never shows up as a Lambda error. That is the right
# behaviour for the person waiting, and exactly why it needs its own alarm:
# otherwise a broken Bedrock setup looks like a series of successful runs that
# quietly describe nothing.
resource "aws_cloudwatch_log_metric_filter" "caption_failures" {
  name           = "${var.function_name}-caption-failures"
  log_group_name = aws_cloudwatch_log_group.lambda.name
  pattern        = "\"Could not generate a caption\""

  metric_transformation {
    name          = "CaptionFailures"
    namespace     = local.metric_namespace
    value         = "1"
    default_value = "0"
    unit          = "Count"
  }
}

resource "aws_cloudwatch_metric_alarm" "caption_failures" {
  alarm_name          = "${var.function_name}-caption-failures"
  alarm_description   = "Pictures are falling back to placeholder alt text instead of being described. Check Bedrock model access and the function's IAM policy."
  namespace           = local.metric_namespace
  metric_name         = aws_cloudwatch_log_metric_filter.caption_failures.metric_transformation[0].name
  statistic           = "Sum"
  period              = 300
  evaluation_periods  = 1
  threshold           = var.caption_failure_threshold
  comparison_operator = "GreaterThanOrEqualToThreshold"

  treat_missing_data = "notBreaching"

  alarm_actions = [aws_sns_topic.alarms.arn]
  ok_actions    = [aws_sns_topic.alarms.arn]
}
